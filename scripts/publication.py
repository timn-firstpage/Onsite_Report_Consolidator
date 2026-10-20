"""Publish validated files without replacing an existing output on either platform."""
import errno
import os
from pathlib import Path
import shutil


def publish_posix(source, output):
    source, output = Path(source), Path(output)
    try:
        os.link(source, output)
        return 'hardlink'
    except FileExistsError:
        raise ValueError(f'output already exists: {output}')
    except OSError as exc:
        unsupported = {errno.EPERM, errno.EXDEV, errno.ENOSYS, errno.ENOTSUP, errno.EOPNOTSUPP}
        if exc.errno not in unsupported:
            raise
    # Some external/FUSE volumes do not support hard links. Exclusive creation
    # still prevents replacement; this fallback is a copy rather than atomic publish.
    owned = False
    try:
        with output.open('xb') as destination:
            owned = True
            with source.open('rb') as original:
                shutil.copyfileobj(original, destination)
            destination.flush()
            os.fsync(destination.fileno())
        return 'exclusive-copy'
    except FileExistsError:
        raise ValueError(f'output already exists: {output}')
    except BaseException:
        if owned:
            output.unlink(missing_ok=True)
        raise


def publish_no_replace(source, output):
    if os.name != 'nt':
        return publish_posix(source, output)
    try:
        os.rename(source, output)
    except FileExistsError:
        raise ValueError(f'output already exists: {output}')
    return 'windows-rename'
