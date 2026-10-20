"""Native CLI and filesystem contracts shared by Windows and macOS runners."""
import errno
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import unittest
from unittest.mock import patch

import openpyxl
import test_consolidator as fixtures

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'scripts/consolidate.py'


class PortabilityTests(unittest.TestCase):
    def setUp(self):
        self.case = fixtures.ConsolidationTests('runTest')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)

    def cli(self, *args):
        return subprocess.run([sys.executable, str(SCRIPT), *map(str, args)],
                              cwd=ROOT / '.runs', capture_output=True, encoding='utf-8')

    def test_doctor_reports_interpreter_and_required_image_support(self):
        result = self.cli('doctor')
        self.assertEqual(result.returncode, 0, result.stderr)
        report = json.loads(result.stdout)
        self.assertTrue(report['ok'])
        self.assertEqual(Path(report['python_executable']).resolve(), Path(sys.executable).resolve())
        self.assertTrue(report['dependencies']['Pillow']['ok'])

    def test_doctor_does_not_write_bytecode_on_fresh_skill_copy(self):
        skill = self.case.folder / 'fresh skill'
        (skill / 'scripts').mkdir(parents=True)
        (skill / 'assets').mkdir()
        for source in (ROOT / 'scripts').iterdir():
            if source.is_file():shutil.copy2(source, skill / 'scripts' / source.name)
        shutil.copy2(ROOT / 'assets/Onsite Checklist - TEMPLATE.xlsx',
                     skill / 'assets/Onsite Checklist - TEMPLATE.xlsx')
        before = sorted(p.relative_to(skill).as_posix() for p in skill.rglob('*'))
        result = subprocess.run([sys.executable, str(skill / 'scripts/consolidate.py'), 'doctor'],
                                cwd=self.case.folder, capture_output=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(sorted(p.relative_to(skill).as_posix() for p in skill.rglob('*')), before)

    def test_relative_backslashes_work_on_both_platforms(self):
        p = self.case.report()
        nested = self.case.folder / 'reports' / '來源'
        nested.mkdir(parents=True)
        p.rename(nested / p.name)
        self.case.manifest['sources'][0]['file'] = 'reports\\來源\\' + p.name
        self.assertEqual(len(self.case.read()['checks']), 9)

    def test_mac_decomposed_company_filename_matches_same_unicode_name(self):
        p = self.case.report()
        name = p.name.replace('Example Co', 'Cafe\u0301 Co')
        p.rename(p.with_name(name))
        self.case.manifest['company'] = 'Café Co'
        self.case.manifest['sources'][0]['file'] = name
        self.assertEqual(len(self.case.read()['checks']), 9)

    def test_quoted_home_path_is_expanded_without_guessing_machine_paths(self):
        from paths import local_path
        self.assertEqual(local_path('~/Documents/audit run/run.json'),
                         (Path.home() / 'Documents/audit run/run.json').resolve())

    @unittest.skipIf(os.name == 'nt', 'native macOS/POSIX path semantics')
    def test_foreign_windows_absolute_path_is_rejected_on_mac(self):
        from paths import local_path
        with self.assertRaisesRegex(ValueError, 'Windows|current machine'):
            local_path(r'C:\Users\Someone\Downloads\run.json')

    @unittest.skipUnless(os.name == 'nt', 'native Windows path semantics')
    def test_foreign_mac_absolute_path_is_rejected_on_windows(self):
        from paths import local_path
        with self.assertRaisesRegex(ValueError, 'POSIX|macOS|current machine'):
            local_path('/Users/someone/Downloads/run.json')

    def test_windows_drive_cannot_be_a_relative_source_on_any_platform(self):
        self.case.report()
        self.case.manifest['sources'][0]['file'] = r'C:\outside.xlsx'
        with self.assertRaisesRegex(ValueError, 'relative|outside'):
            self.case.read()

    def test_posix_publication_falls_back_without_hardlink_support(self):
        from publication import publish_posix
        source = self.case.folder / 'validated temp.xlsx'
        output = self.case.folder / '最終 report.xlsx'
        source.write_bytes(b'complete validated workbook')
        with patch('publication.os.link', side_effect=OSError(errno.ENOTSUP, 'unsupported')):
            publish_posix(source, output)
        self.assertEqual(output.read_bytes(), b'complete validated workbook')
        with patch('publication.os.link', side_effect=OSError(errno.ENOTSUP, 'unsupported')):
            with self.assertRaisesRegex(ValueError, 'exists'):
                publish_posix(source, output)
        self.assertEqual(output.read_bytes(), b'complete validated workbook')

    def test_failed_posix_fallback_cleans_only_its_own_partial_output(self):
        from publication import publish_posix
        source = self.case.folder / 'validated.xlsx'
        output = self.case.folder / 'final.xlsx'
        source.write_bytes(b'validated input')
        with patch('publication.os.link', side_effect=OSError(errno.ENOTSUP, 'unsupported')):
            with patch('publication.shutil.copyfileobj', side_effect=OSError('disk full')):
                with self.assertRaisesRegex(OSError, 'disk full'):
                    publish_posix(source, output)
        self.assertFalse(output.exists())
        self.assertEqual(source.read_bytes(), b'validated input')

    def image_template(self):
        from PIL import Image
        from openpyxl.drawing.image import Image as ExcelImage
        png = self.case.folder / 'template image.png'
        Image.new('RGB', (8, 8), 'navy').save(png)
        w = openpyxl.load_workbook(self.case.template)
        w['11. Sitemap'].add_image(ExcelImage(str(png)), 'D4')
        w.save(self.case.template)

    def test_cli_build_verify_unicode_spaces_and_template_image_from_other_cwd(self):
        folder = self.case.folder / '客戶 café run with spaces'
        folder.mkdir()
        shutil.move(self.case.template, folder / 'template.xlsx')
        self.case.folder = folder
        self.case.path = folder / 'run.json'
        self.case.template = folder / 'template.xlsx'
        self.image_template()
        p = self.case.report(statuses={'11.3': 'Human check'})
        renamed = p.with_name(p.name.replace('Example Co', '客戶 Café Co'))
        p.rename(renamed)
        self.case.manifest['company'] = '客戶 Café Co'
        self.case.manifest['sources'][0]['file'] = renamed.name
        self.case.save()
        inspected = self.cli('inspect', renamed)
        self.assertEqual(inspected.returncode, 0, inspected.stderr)
        result = self.cli('build', '--manifest', self.case.path)
        self.assertEqual(result.returncode, 0, result.stderr)
        output = Path(json.loads(result.stdout)['output'])
        self.assertEqual(output.parent, folder)
        w = openpyxl.load_workbook(output)
        self.assertEqual(w.active.title, '(auto) SEO Checklist')
        self.assertEqual(w.active['D8'].value, 'Human check')
        self.assertEqual(len(w['11. Sitemap']._images), 1)
        w.close()
        result = self.cli('verify', output)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)['checks'], 9)
        original = output.read_bytes()
        collision = self.cli('build', '--manifest', self.case.path)
        self.assertEqual(collision.returncode, 2, collision.stderr)
        self.assertEqual(original, output.read_bytes())

    def test_missing_pillow_stops_before_silently_dropping_template_images(self):
        self.image_template()
        self.case.report()
        manifest = self.case.save()
        code = '''import importlib.abc,sys
class BlockPillow(importlib.abc.MetaPathFinder):
 def find_spec(self,fullname,path=None,target=None):
  if fullname=='PIL' or fullname.startswith('PIL.'):
   raise ModuleNotFoundError('Pillow deliberately unavailable')
sys.meta_path.insert(0,BlockPillow())
sys.path.insert(0,sys.argv[1])
sys.argv=['consolidate.py','build','--manifest',sys.argv[2]]
import consolidate
raise SystemExit(consolidate.main())
'''
        result = subprocess.run([sys.executable, '-c', code, str(ROOT / 'scripts'), str(manifest)],
                                cwd=ROOT / '.runs', capture_output=True, encoding='utf-8')
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn('Pillow', result.stderr)
        self.assertFalse(list(self.case.folder.glob('Autocheck*.xlsx')))


if __name__ == '__main__':
    unittest.main()
