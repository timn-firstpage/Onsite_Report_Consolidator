"""Read-only environment checks; no installation, credential or network operations."""
import importlib
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
import platform
import re
import sys

DEPENDENCIES = {
    'openpyxl': ('openpyxl', (3, 1, 5), (4,)),
    'Pillow': ('PIL.Image', (10, 4), (14,)),
}


def runtime_report():
    dependencies = {}
    for package, (module, minimum, upper) in DEPENDENCIES.items():
        try:
            importlib.import_module(module)
            installed = version(package)
            numeric = tuple(int(x) for x in re.findall(r'\d+', installed)[:3])
            ok = minimum <= numeric < upper
            dependencies[package] = {'ok': ok, 'version': installed}
            if not ok:
                dependencies[package]['error'] = 'version outside requirements.txt bounds'
        except (ImportError, PackageNotFoundError, OSError) as exc:
            dependencies[package] = {'ok': False, 'version': None, 'error': str(exc)}
    template = Path(__file__).resolve().parents[1] / 'assets/Onsite Checklist - TEMPLATE.xlsx'
    python_ok = sys.version_info >= (3, 10)
    return {'ok': python_ok and all(d['ok'] for d in dependencies.values()) and template.is_file(),
            'platform': platform.system(), 'architecture': platform.machine(),
            'python_version': platform.python_version(), 'python_executable': sys.executable,
            'python_supported': python_ok, 'dependencies': dependencies,
            'bundled_template_exists': template.is_file()}


def require_runtime():
    report = runtime_report()
    if not report['python_supported']:
        raise ValueError('Python 3.10 or newer is required')
    missing = [name for name, info in report['dependencies'].items() if not info['ok']]
    if missing:
        raise ValueError('Missing/incompatible dependency: ' + ', '.join(missing) +
                         '. Install requirements.txt using this same Python interpreter; '
                         'Pillow is required to preserve template images.')
    return report
