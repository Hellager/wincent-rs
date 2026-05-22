use std::collections::HashSet;

pub(super) const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
pub(super) const FREE_SECTOR: u32 = 0xFFFF_FFFF;
const FAT_SECTOR: u32 = 0xFFFF_FFFD;
const DIFAT_SECTOR: u32 = 0xFFFF_FFFC;

pub(super) type CfbResult<T> = Result<T, String>;

#[derive(Debug, Clone)]
pub(super) struct DirectoryEntry {
    pub(super) name: String,
    pub(super) object_type: u8,
    pub(super) start_sector: u32,
    pub(super) stream_size: u64,
}

#[derive(Debug)]
pub(super) struct CompoundFile {
    pub(super) data: Vec<u8>,
    pub(super) sector_size: usize,
    pub(super) mini_sector_size: usize,
    pub(super) mini_cutoff_size: u32,
    pub(super) fat: Vec<u32>,
    mini_fat: Vec<u32>,
    pub(super) directory: Vec<DirectoryEntry>,
    root_stream: Vec<u8>,
}

impl CompoundFile {
    pub(super) fn parse(data: Vec<u8>) -> CfbResult<Self> {
        if data.len() < 512 {
            return Err("file is too small for a CFB header".to_string());
        }
        if data.get(0..8) != Some(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) {
            return Err("not an OLE Compound File Binary".to_string());
        }

        let sector_size = 1usize
            .checked_shl(read_u16(&data, 0x1e)? as u32)
            .ok_or_else(|| "invalid sector size shift".to_string())?;
        if sector_size != 512 && sector_size != 4096 {
            return Err(format!(
                "unsupported CFB sector size {sector_size}: expected 512 or 4096"
            ));
        }
        let mini_sector_size = 1usize
            .checked_shl(read_u16(&data, 0x20)? as u32)
            .ok_or_else(|| "invalid mini sector size shift".to_string())?;
        if mini_sector_size != 64 {
            return Err(format!(
                "unsupported CFB mini sector size {mini_sector_size}: expected 64"
            ));
        }
        let first_dir_sector = read_u32(&data, 0x30)?;
        let mini_cutoff_size = read_u32(&data, 0x38)?;
        let first_mini_fat_sector = read_u32(&data, 0x3c)?;
        let num_mini_fat_sectors = read_u32(&data, 0x40)? as usize;
        let first_difat_sector = read_u32(&data, 0x44)?;
        let num_difat_sectors = read_u32(&data, 0x48)? as usize;

        let fat_sector_ids = read_difat(&data, sector_size, first_difat_sector, num_difat_sectors)?;
        let fat = read_fat(&data, sector_size, &fat_sector_ids)?;

        let directory_stream = read_regular_stream(&data, sector_size, &fat, first_dir_sector)?;
        let directory = parse_directory(&directory_stream)?;
        let root = directory
            .iter()
            .find(|entry| entry.object_type == 5)
            .ok_or_else(|| "root storage entry not found".to_string())?;
        let root_stream = read_regular_stream_sized(
            &data,
            sector_size,
            &fat,
            root.start_sector,
            root.stream_size as usize,
        )?;

        let mini_fat = if first_mini_fat_sector == FREE_SECTOR || num_mini_fat_sectors == 0 {
            Vec::new()
        } else {
            let bytes = read_regular_stream_sized(
                &data,
                sector_size,
                &fat,
                first_mini_fat_sector,
                num_mini_fat_sectors * sector_size,
            )?;
            bytes
                .chunks_exact(4)
                .map(|chunk| u32::from_le_bytes(chunk.try_into().expect("4-byte chunk")))
                .collect()
        };

        Ok(Self {
            data,
            sector_size,
            mini_sector_size,
            mini_cutoff_size,
            fat,
            mini_fat,
            directory,
            root_stream,
        })
    }

    pub(super) fn stream(&self, name: &str) -> CfbResult<Option<Vec<u8>>> {
        let Some(entry) = self
            .directory
            .iter()
            .find(|entry| entry.object_type == 2 && entry.name.eq_ignore_ascii_case(name))
        else {
            return Ok(None);
        };

        let size = entry.stream_size as usize;
        if size < self.mini_cutoff_size as usize {
            self.read_mini_stream(entry.start_sector, size).map(Some)
        } else {
            read_regular_stream_sized(
                &self.data,
                self.sector_size,
                &self.fat,
                entry.start_sector,
                size,
            )
            .map(Some)
        }
    }

    fn read_mini_stream(&self, start_sector: u32, size: usize) -> CfbResult<Vec<u8>> {
        if size == 0 {
            return Ok(Vec::new());
        }

        let chain = sector_chain(&self.mini_fat, start_sector)?;
        let mut out = Vec::with_capacity(size);
        for mini_sector in chain {
            let offset = mini_sector as usize * self.mini_sector_size;
            let end = offset + self.mini_sector_size;
            let sector = self
                .root_stream
                .get(offset..end)
                .ok_or_else(|| format!("mini sector {mini_sector} is out of bounds"))?;
            out.extend_from_slice(sector);
            if out.len() >= size {
                break;
            }
        }
        out.truncate(size);
        Ok(out)
    }
}

pub(super) fn sector_chain(fat: &[u32], start_sector: u32) -> CfbResult<Vec<u32>> {
    if start_sector == FREE_SECTOR || start_sector == END_OF_CHAIN {
        return Ok(Vec::new());
    }

    let mut chain = Vec::new();
    let mut seen = HashSet::new();
    let mut sector = start_sector;

    loop {
        if sector == END_OF_CHAIN {
            break;
        }
        if sector == FREE_SECTOR || sector == FAT_SECTOR || sector == DIFAT_SECTOR {
            return Err(format!("invalid sector marker {sector:#x} in chain"));
        }
        let index = sector as usize;
        if index >= fat.len() {
            return Err(format!("sector {sector} is outside the FAT"));
        }
        if !seen.insert(sector) {
            return Err(format!("loop detected in sector chain at sector {sector}"));
        }
        chain.push(sector);
        sector = fat[index];
    }

    Ok(chain)
}

pub(super) fn read_regular_stream_sized(
    data: &[u8],
    sector_size: usize,
    fat: &[u32],
    start_sector: u32,
    size: usize,
) -> CfbResult<Vec<u8>> {
    let mut out = read_regular_stream(data, sector_size, fat, start_sector)?;
    out.truncate(size);
    Ok(out)
}

fn read_regular_stream(
    data: &[u8],
    sector_size: usize,
    fat: &[u32],
    start_sector: u32,
) -> CfbResult<Vec<u8>> {
    let chain = sector_chain(fat, start_sector)?;
    let mut out = Vec::new();
    for sector_id in chain {
        out.extend_from_slice(sector_slice(data, sector_size, sector_id)?);
    }
    Ok(out)
}

fn read_difat(
    data: &[u8],
    sector_size: usize,
    first_difat_sector: u32,
    num_difat_sectors: usize,
) -> CfbResult<Vec<u32>> {
    let mut fat_sector_ids = Vec::new();
    for index in 0..109 {
        let sector_id = read_u32(data, 0x4c + index * 4)?;
        if sector_id != FREE_SECTOR {
            fat_sector_ids.push(sector_id);
        }
    }

    let mut next_difat_sector = first_difat_sector;
    for _ in 0..num_difat_sectors {
        if next_difat_sector == FREE_SECTOR || next_difat_sector == END_OF_CHAIN {
            break;
        }

        let sector = sector_slice(data, sector_size, next_difat_sector)?;
        let entries_per_sector = sector_size / 4;
        for index in 0..entries_per_sector - 1 {
            let sector_id = u32::from_le_bytes(
                sector[index * 4..index * 4 + 4]
                    .try_into()
                    .expect("4-byte chunk"),
            );
            if sector_id != FREE_SECTOR {
                fat_sector_ids.push(sector_id);
            }
        }
        next_difat_sector = u32::from_le_bytes(
            sector[(entries_per_sector - 1) * 4..entries_per_sector * 4]
                .try_into()
                .expect("4-byte chunk"),
        );
    }

    Ok(fat_sector_ids)
}

fn read_fat(data: &[u8], sector_size: usize, fat_sector_ids: &[u32]) -> CfbResult<Vec<u32>> {
    let mut fat = Vec::new();
    for sector_id in fat_sector_ids {
        let sector = sector_slice(data, sector_size, *sector_id)?;
        fat.extend(
            sector
                .chunks_exact(4)
                .map(|chunk| u32::from_le_bytes(chunk.try_into().expect("4-byte chunk"))),
        );
    }
    Ok(fat)
}

fn sector_slice(data: &[u8], sector_size: usize, sector_id: u32) -> CfbResult<&[u8]> {
    let offset = (sector_id as usize + 1)
        .checked_mul(sector_size)
        .ok_or_else(|| "sector offset overflow".to_string())?;
    let end = offset
        .checked_add(sector_size)
        .ok_or_else(|| "sector end overflow".to_string())?;
    data.get(offset..end)
        .ok_or_else(|| format!("sector {sector_id} is out of bounds"))
}

fn parse_directory(directory_stream: &[u8]) -> CfbResult<Vec<DirectoryEntry>> {
    let mut entries = Vec::new();
    for chunk in directory_stream.chunks_exact(128) {
        let object_type = chunk[66];
        if object_type == 0 {
            continue;
        }

        let name_len = u16::from_le_bytes(chunk[64..66].try_into().expect("2-byte chunk"));
        let name = if name_len >= 2 && name_len as usize <= 64 {
            decode_utf16_lossy(&chunk[..name_len as usize - 2])
        } else {
            String::new()
        };

        entries.push(DirectoryEntry {
            name,
            object_type,
            start_sector: u32::from_le_bytes(chunk[116..120].try_into().expect("4-byte chunk")),
            stream_size: u64::from_le_bytes(chunk[120..128].try_into().expect("8-byte chunk")),
        });
    }
    Ok(entries)
}

pub(super) fn decode_utf16_lossy(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes(chunk.try_into().expect("2-byte chunk")))
        .collect();
    String::from_utf16_lossy(&units)
}

pub(super) fn read_u16(data: &[u8], offset: usize) -> CfbResult<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(u16::from_le_bytes(bytes.try_into().expect("2-byte chunk")))
}

pub(super) fn read_u32(data: &[u8], offset: usize) -> CfbResult<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(u32::from_le_bytes(bytes.try_into().expect("4-byte chunk")))
}

pub(super) fn read_i32(data: &[u8], offset: usize) -> CfbResult<i32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(i32::from_le_bytes(bytes.try_into().expect("4-byte chunk")))
}

pub(super) fn read_u64(data: &[u8], offset: usize) -> CfbResult<u64> {
    let bytes = data
        .get(offset..offset + 8)
        .ok_or_else(|| format!("unexpected end of data at offset {offset}"))?;
    Ok(u64::from_le_bytes(bytes.try_into().expect("8-byte chunk")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_wrong_magic() {
        let mut data = vec![0u8; 512];
        // Wrong magic bytes
        data[0..8].copy_from_slice(&[0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
        let result = CompoundFile::parse(data);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("OLE Compound File Binary"));
    }

    #[test]
    fn rejects_too_small() {
        let data = vec![0u8; 100];
        let result = CompoundFile::parse(data);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("too small"));
    }

    #[test]
    fn sector_chain_detects_loop() {
        // fat[0] -> 1, fat[1] -> 0 (loop)
        let fat = vec![1u32, 0u32];
        let result = sector_chain(&fat, 0);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("loop detected"));
    }

    #[test]
    fn sector_chain_empty_for_free_sector() {
        let fat = vec![END_OF_CHAIN];
        let result = sector_chain(&fat, FREE_SECTOR);
        assert!(result.is_ok());
        assert!(result.unwrap().is_empty());
    }

    #[test]
    fn sector_chain_out_of_bounds() {
        let fat = vec![END_OF_CHAIN];
        // sector 5 is outside fat of length 1
        let result = sector_chain(&fat, 5);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("outside the FAT"));
    }

    #[test]
    fn rejects_invalid_sector_size() {
        let mut data = vec![0u8; 512];
        data[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        // Sector size shift of 1 → sector_size = 2, which is neither 512 nor 4096.
        data[0x1e..0x20].copy_from_slice(&1u16.to_le_bytes());
        data[0x20..0x22].copy_from_slice(&6u16.to_le_bytes()); // mini: 64
        let result = CompoundFile::parse(data);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("unsupported CFB sector size"));
    }
}
