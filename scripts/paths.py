"""Native machine paths and portable relative manifest paths."""
import os
from pathlib import Path, PurePosixPath, PureWindowsPath


def local_path(value, base=None):
    text = os.fspath(value)
    if os.name != 'nt' and PureWindowsPath(text).drive:
        raise ValueError(f'Windows path cannot be used on the current machine: {text}')
    if os.name == 'nt' and text.startswith('/') and not text.startswith('//'):
        raise ValueError(f'POSIX/macOS path cannot be used on the current machine: {text}')
    path = Path(text).expanduser()
    if base is not None and not path.is_absolute():
        path = Path(base) / path
    return path.resolve()


def relative_input(folder, name):
    if not isinstance(name, str):
        raise ValueError('source file must be a relative path string')
    windows = PureWindowsPath(name)
    portable = PurePosixPath(name.replace('\\', '/'))
    if windows.drive or windows.root or portable.is_absolute():
        raise ValueError(f'file must be relative and inside run folder: {name}')
    resolved = (folder / Path(*portable.parts)).resolve()
    if not resolved.is_relative_to(folder) or resolved == folder:
        raise ValueError(f'file must be relative and inside run folder: {name}')
    if not resolved.is_file():
        raise ValueError(f'missing input file: {name}')
    return resolved
