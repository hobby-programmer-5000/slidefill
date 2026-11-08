"""python -m verification generate INPUT.xlsx NEW_FOLDER --lang en|ko
python -m verification check INPUT.xlsx REPORT.pptx CERTIFICATE.json
python -m verification selftest
"""
import argparse
import json
import sys


def main():
    parser=argparse.ArgumentParser(description='CPU-only mathematical/source/report verification; no local model.')
    sub=parser.add_subparsers(dest='command',required=True)
    g=sub.add_parser('generate');g.add_argument('workbook');g.add_argument('output');g.add_argument('--lang',choices=['en','ko'],default='en')
    c=sub.add_parser('check');c.add_argument('workbook');c.add_argument('report');c.add_argument('certificate')
    sub.add_parser('selftest')
    args=parser.parse_args()
    if args.command=='selftest':
        import unittest
        from . import test_verification,test_automatic,test_coverage,test_binary,test_remaining,test_legacy_html,test_formulas,test_reliability,test_standalone_formulas
        suite=unittest.TestSuite(unittest.defaultTestLoader.loadTestsFromModule(module) for module in (test_verification,test_automatic,test_coverage,test_binary,test_remaining,test_legacy_html,test_formulas,test_reliability,test_standalone_formulas))
        result=unittest.TextTestRunner(verbosity=2).run(suite)
        return 0 if result.wasSuccessful() else 1
    try:
        if args.command=='generate':
            from .pipeline import generate
            result=generate(args.workbook,args.output,args.lang)
        else:
            from .check import check
            with open(args.certificate,encoding='utf-8') as f:certificate=json.load(f)
            result=check(args.workbook,args.report,certificate)
        print(json.dumps(result,ensure_ascii=False,indent=2))
        return 0 if result['result']=='PASS' else 1
    except (ValueError,OSError) as exc:
        print(f'Cannot verify: {exc}',file=sys.stderr)
        return 2


if __name__=='__main__':
    raise SystemExit(main())
