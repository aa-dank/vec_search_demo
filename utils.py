from pathlib import Path, PurePosixPath


def extract_server_dirs(full_path: str | Path,
                       base_mount: str | Path,
                       include_filename: bool = False) -> str:
    """
    Extract the server directory path (and optionally filename) relative to a mount point.

    Parameters
    ----------
    full_path   Absolute path on the client machine
                e.g. r"N:\\PPDO\\Records\\49xx   Long Marine Lab\\4932\\file.pdf"
    base_mount  The local mount-point for the records share
                e.g. r"N:\\PPDO\\Records"   or   "/mnt/records"
    include_filename : bool, default False
                Whether to include the filename in the returned path.
                If False, only returns the directory structure.

    Returns
    -------
    str   --  value suitable for file_locations.file_server_directories
              (always forward-slash separators, no leading slash)
              If include_filename=False, excludes the filename component.
    """
    # Normalise to platform-aware Path objects
    full = Path(full_path).expanduser().resolve()
    base = Path(base_mount).expanduser().resolve()

    # 1) Get the sub-path *relative* to the mount
    try:
        rel_parts = full.relative_to(base)
    except ValueError:               # not under base_mount
        raise ValueError(f"{full} is not under {base}")

    # 2) If include_filename is False, exclude the filename (last part)
    if not include_filename and rel_parts.parts:
        # Remove the last part (filename) if it exists
        rel_parts = Path(*rel_parts.parts[:-1]) if len(rel_parts.parts) > 1 else Path()

    # 3) Convert to POSIX form (forces forward slashes)
    return str(PurePosixPath(rel_parts))

def build_file_path(base_mount: str,
                    server_dir: str,
                    filename: str = None) -> Path:
    """
    Join a server-relative path + filename onto a machine-specific
    mount-point.

    Parameters
    ----------
    base_mount : str
        The local mount of the records share, e.g.
        r"N:\PPDO\Records"  (Windows)  or  "/mnt/records" (Linux).
    server_dir : str
        The value from file_locations.file_server_directories
        (always stored with forward-slashes).
    filename   : str
        file_locations.filename

    Returns
    -------
    pathlib.Path  – ready for open(), exists(), etc.
    """
    # 1) Treat the DB field as a *POSIX* path (it always uses “/”)
    rel_parts = PurePosixPath(server_dir).parts     # -> tuple of segments

    # 2) Let Path figure out the separator style of this machine
    full_path = Path(base_mount).joinpath(*rel_parts)
    if filename:
        full_path = full_path / filename
    
    return full_path