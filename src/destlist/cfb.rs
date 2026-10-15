use std::io::{Cursor, Read};
use std::path::PathBuf;

#[cfg(test)]
use std::path::Path;

pub(super) type CfbResult<T> = Result<T, String>;

#[derive(Debug, Clone)]
pub(super) struct DirectoryEntry {
    pub(super) name: String,
    pub(super) path: PathBuf,
    pub(super) object_type: u8,
    pub(super) start_sector: Option<u32>,
    pub(super) stream_size: u64,
}

#[derive(Debug, Clone)]
pub(super) struct StreamData {
    pub(super) name: String,
    pub(super) path: PathBuf,
    pub(super) data: Vec<u8>,
}

#[derive(Debug)]
pub(super) struct CompoundFile {
    pub(super) major_version: u16,
    pub(super) sector_size: usize,
    pub(super) mini_sector_size: usize,
    pub(super) mini_cutoff_size: u32,
    pub(super) directory: Vec<DirectoryEntry>,
    pub(super) streams: Vec<StreamData>,
}

impl CompoundFile {
    pub(super) fn parse(data: Vec<u8>) -> CfbResult<Self> {
        let mut compound = ::cfb::CompoundFile::open(Cursor::new(data)).map_err(|error| {
            format!("failed to parse OLE Compound File Binary container: {error}")
        })?;

        let (major_version, sector_size) = match compound.version() {
            ::cfb::Version::V3 => (3, 512),
            ::cfb::Version::V4 => (4, 4096),
        };

        let entries: Vec<_> = compound.walk().collect();
        let directory = entries
            .iter()
            .map(|entry| DirectoryEntry {
                name: entry.name().to_string(),
                path: entry.path().to_path_buf(),
                object_type: object_type(entry),
                // rust-cfb exposes logical objects, not allocator-specific sectors.
                start_sector: None,
                stream_size: entry.len(),
            })
            .collect();

        let mut streams = Vec::new();
        for entry in entries.iter().filter(|entry| entry.is_stream()) {
            streams.push(read_stream(&mut compound, entry)?);
        }

        Ok(Self {
            major_version,
            sector_size,
            mini_sector_size: 64,
            mini_cutoff_size: 4096,
            directory,
            streams,
        })
    }

    pub(super) fn stream(&self, name: &str) -> Option<&[u8]> {
        self.streams
            .iter()
            .find(|stream| stream.name.eq_ignore_ascii_case(name))
            .map(|stream| stream.data.as_slice())
    }

    #[cfg(test)]
    pub(super) fn stream_at(&self, path: &Path) -> Option<&[u8]> {
        self.streams
            .iter()
            .find(|stream| stream.path == path)
            .map(|stream| stream.data.as_slice())
    }
}

fn object_type(entry: &::cfb::Entry) -> u8 {
    if entry.is_root() {
        5
    } else if entry.is_stream() {
        2
    } else {
        1
    }
}

fn read_stream(
    compound: &mut ::cfb::CompoundFile<Cursor<Vec<u8>>>,
    entry: &::cfb::Entry,
) -> CfbResult<StreamData> {
    let expected_len = usize::try_from(entry.len()).map_err(|_| {
        format!(
            "CFB stream `{}` is too large to fit in memory: {} bytes",
            entry.path().display(),
            entry.len()
        )
    })?;
    let mut stream = compound.open_stream(entry.path()).map_err(|error| {
        format!(
            "failed to open CFB stream `{}`: {error}",
            entry.path().display()
        )
    })?;
    let mut data = Vec::with_capacity(expected_len);
    stream.read_to_end(&mut data).map_err(|error| {
        format!(
            "failed to read CFB stream `{}`: {error}",
            entry.path().display()
        )
    })?;
    if data.len() != expected_len {
        return Err(format!(
            "CFB stream `{}` declared {expected_len} bytes but yielded {} bytes",
            entry.path().display(),
            data.len()
        ));
    }

    Ok(StreamData {
        name: entry.name().to_string(),
        path: entry.path().to_path_buf(),
        data,
    })
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
    use std::io::Write;

    #[test]
    fn rejects_wrong_magic() {
        let data = vec![0u8; 512];
        let result = CompoundFile::parse(data);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("OLE Compound File Binary"));
    }

    #[test]
    fn rejects_too_small() {
        let data = vec![0u8; 100];
        let result = CompoundFile::parse(data);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("OLE Compound File Binary"));
    }

    #[test]
    fn reads_every_stream_from_v3_container() {
        let cursor = Cursor::new(Vec::new());
        let mut writer = ::cfb::CompoundFile::create_with_version(::cfb::Version::V3, cursor)
            .expect("failed to create test CFB");

        writer
            .create_stream("/DestList")
            .expect("failed to create DestList")
            .write_all(b"dest-list")
            .expect("failed to write DestList");
        writer
            .create_stream("/1")
            .expect("failed to create numeric stream")
            .write_all(b"shell-link")
            .expect("failed to write numeric stream");
        writer.flush().expect("failed to flush test CFB");

        let parsed = CompoundFile::parse(writer.into_inner().into_inner())
            .expect("failed to parse test CFB");

        assert_eq!(parsed.major_version, 3);
        assert_eq!(parsed.streams.len(), 2);
        assert_eq!(parsed.stream("DestList"), Some(b"dest-list".as_slice()));
        assert_eq!(parsed.stream("1"), Some(b"shell-link".as_slice()));
        assert_eq!(
            parsed.stream_at(Path::new("/DestList")),
            Some(b"dest-list".as_slice())
        );
    }
}
