#!/usr/bin/env python3
"""Package a generated report for Pages. Publish only the report and public inputs."""
import hashlib
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def main():
    report=json.loads((ROOT/'data/processed/report.json').read_text())
    if report['refresh_failures']:
        raise ValueError('A source update failed. Keep the last published report.')
    content=(ROOT/'output/index.html').read_text()
    expected='href="../${esc(s.path)}"'
    if expected not in content:
        raise ValueError('Check the source file links before publishing.')
    content=content.replace(expected,'href="${esc(s.path)}"')
    if '/Users/' in content:
        raise ValueError('The report contains a local computer path.')
    target=ROOT/'_site'; target.mkdir(exist_ok=True)
    for dest in [target,ROOT]:
        (dest/'index.html').write_text(content)
        shutil.copy2(ROOT/'output/series.csv',dest/'series.csv')
    (target/'.nojekyll').write_text('')
    for source in report['sources'].values():
        relative=Path(source['path'])
        if relative.is_absolute() or '..' in relative.parts or relative.parts[:2]!=('data','raw'):
            raise ValueError('Unexpected source file path.')
        original=ROOT/relative
        if hashlib.sha256(original.read_bytes()).hexdigest()!=source['sha256']:
            raise ValueError('Source hash mismatch: '+source['id'])
        destination=target/relative
        destination.parent.mkdir(parents=True,exist_ok=True)
        shutil.copy2(original,destination)
    print('Packaged '+str(len(report['cards']))+' charts and '+str(len(report['sources']))+' source files for GitHub Pages.')

if __name__=='__main__': main()
