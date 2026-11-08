"""Rebuild the shipped integer-only kernel with pinned Dafny4.11.0.

Usage: python verification/proofs/build.py PATH_TO_DAFNY NEW_OUTPUT_FOLDER
The caller must preserve full verifier stdout/exit status. No no-verify/axioms.
"""
import ast
import hashlib
import json
from pathlib import Path
import subprocess
import sys


def extract(path):
    source=ast.parse(path.read_text(encoding='utf-8-sig'))
    classes=[n for n in source.body if isinstance(n,ast.ClassDef)]
    if len(classes)!=1 or classes[0].name!='default__':raise ValueError('unexpected generated class')
    cls=classes[0]
    if cls.bases or cls.keywords:raise ValueError('unsupported generated inheritance')
    for fn in cls.body:
        if not isinstance(fn,ast.FunctionDef):raise ValueError('unexpected generated declaration')
        locals_={a.arg for a in fn.args.args}|{n.id for n in ast.walk(fn) if isinstance(n,ast.Name) and isinstance(n.ctx,ast.Store)}
        external={n.id for n in ast.walk(fn) if isinstance(n,ast.Name) and isinstance(n.ctx,ast.Load)}-locals_
        if external-{'staticmethod','int','bool','len'}:raise ValueError(f'runtime-dependent generated code: {external}')
    # Only unused module imports are removed. The compiled class AST is retained.
    result='# Generated from verification/proofs/Kernel.dfy by Dafny4.11.0.\n# MIT, original project code. Rebuild with proofs/build.py; do not edit.\n'+ast.unparse(cls)+'\n'
    assert ast.dump(ast.parse(result).body[0])==ast.dump(cls)
    return result


def build(executable,output):
    root=Path(__file__).resolve().parents[1];proof=root/'proofs/Kernel.dfy'
    out=Path(output).resolve();out.mkdir(parents=True,exist_ok=False)
    version=subprocess.run([executable,'--version'],capture_output=True,text=True,check=True).stdout.strip()
    if not version.startswith('4.11.0+'):raise ValueError('requires pinned Dafny4.11.0')
    source=proof.read_text(encoding='utf-8')
    code='\n'.join(line.split('//')[0] for line in source.splitlines())
    for forbidden in ('assume','axiom','extern','verify false','{:verify'):
        if forbidden in code:raise ValueError('untrusted proof directive')
    command=[executable,'build','--target:py','--spill-translation','--enforce-determinism','--verification-time-limit','20','--output',str(out/'Kernel'),str(proof)]
    result=subprocess.run(command,capture_output=True,text=True)
    (out/'verify.log').write_text(result.stdout+result.stderr,encoding='utf-8')
    print(result.stdout+result.stderr)
    if result.returncode or '9 verified, 0 errors' not in result.stdout:raise RuntimeError('proof did not verify')
    generated=out/'Kernel-py/VerifiedKernel.py'
    text=extract(generated)
    (out/'verified_kernel.py').write_text(text,encoding='utf-8',newline='\n')
    receipt={'dafny':version,'source_sha256':hashlib.sha256(proof.read_bytes()).hexdigest(),
             'generated_class_sha256':hashlib.sha256(text.encode()).hexdigest(),'verified':9,'errors':0,
             'transformation':'unused imports removed, generated class AST identical','runtime':'integer-only subset; no Dafny runtime copied',
             'trust':'Dafny/Boogie/Z3, Python integers, compiler, extraction script, bridge conversions'}
    (out/'receipt.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
    return out


if __name__=='__main__':build(sys.argv[1],sys.argv[2])
