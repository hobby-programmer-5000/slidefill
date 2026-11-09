import tempfile
import unittest
from fractions import Fraction
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

import core
import native_vba_cases
from verification.rawxlsx import Sheet


class NativeVbaCases(unittest.TestCase):
    def test_opt_in_collects_real_aggregates_without_changing_report(self):
        with tempfile.TemporaryDirectory() as folder:
            book_path, template = write_comparison_fixture(folder)
            normal = core.run(template, book_path)
            compared = core.run(template, book_path, collect_vba=True)

        self.assertIsNone(normal.vba_cases)
        self.assertEqual(compared.verification['state'], 'passed', compared.verification)
        self.assertEqual(compared.verification, normal.verification)
        self.assertEqual(compared.items, normal.items)
        self.assertEqual(compared.vba_cases.skipped, 0)
        cases = compared.vba_cases.cases
        self.assertEqual([case['operation'] for case in cases], ['sum', 'mean', 'count'])
        self.assertEqual([case['expected'] for case in cases], [Fraction(3, 10), Fraction(3, 20), Fraction(2)])
        for case in cases:
            self.assertEqual(case['operands'], [Fraction(1, 10), Fraction(1, 5), None])
            self.assertIn(f"latest1:Data!A,{case['operation']}(B)", case['label'])

    def test_incompatible_operand_types_and_semantics_are_skipped(self):
        sheet = Sheet(
            {'A1': '1', 'A2': '12', 'A3': True, 'A4': '9007199254740993', 'A5': '2', 'A6': '3', 'A7': '4'},
            [], {'A5': 'yyyy-mm-dd'},
            types={'A1': 'n', 'A2': 's', 'A3': 'b', 'A4': 'n', 'A5': 'n', 'A6': 'n', 'A7': 'n'},
            formulas={'A6': {}}, formula_regions=['A7:A9'],
        )
        bindings = [_binding(_aggregate(f'A{row}')) for row in range(1, 9)]
        bindings.append(_binding(_aggregate('A1', filters=[{'wanted': '2025'}])))
        bindings.append(_binding(_aggregate('A1', 'A2')))
        bindings.append(_binding(_aggregate('A1', kind='rate')))
        bindings.append(_binding({'op': 'ratio'}))
        bindings.append(_binding(_aggregate('A1'), path='data.csv'))

        found = native_vba_cases.collect(bindings, {'data.xlsx': {'Data': sheet}, 'data.csv': {'Data': sheet}})

        self.assertEqual(found.skipped, 11)
        self.assertEqual(len(found.cases), 1)
        self.assertEqual(found.cases[0]['expected'], Fraction(1))

    def test_case_cap_is_reported(self):
        sheet = Sheet({'A1': '1'}, [], {}, types={'A1': 'n'})
        bindings = [_binding(_aggregate('A1')) for _ in range(2)]
        with patch.object(native_vba_cases, 'MAX_CASES', 1):
            found = native_vba_cases.collect(bindings, {'data.xlsx': {'Data': sheet}})
        self.assertEqual(len(found.cases), 1)
        self.assertEqual(found.skipped, 1)


def write_comparison_fixture(folder):
    book_path = Path(folder) / 'data.xlsx'
    template = Path(folder) / 'template.pptx'
    book = openpyxl.Workbook()
    book.active.title = 'Data'
    for row in [('Day', 'Amount'), ('2024-01-01', 0.1), ('2024-01-01', 0.2), ('2024-01-01', None)]:
        book.active.append(row)
    book.save(book_path)

    presentation = Presentation()
    for operation in ('sum', 'mean', 'count'):
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        if operation == 'count':
            kind = 'table'
            shape = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(8), Inches(1))
        else:
            kind = 'chart'
            data = CategoryChartData()
            data.categories = ['Old']
            data.add_series('Old', [1])
            shape = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(8), Inches(4), data,
            )
        shape._element.xpath('.//p:cNvPr')[0].set('descr', f'{{{{{kind}:latest1:Data!A,{operation}(B)}}}}')
    presentation.save(template)
    return book_path, template


def _aggregate(*refs, kind='sum', filters=()):
    return {
        'op': 'aggregate', 'kind': kind, 'product': False,
        'rows': [{'cells': [{'op': 'cell', 'sheet': 'Data', 'cell': ref} for ref in refs],
                  'filters': filters, 'numerator': None}],
    }


def _binding(expr, path='data.xlsx'):
    return {
        'kind': 'table', 'key': 'latest1:Data!B,sum(A)', 'part': 'ppt/slides/slide1.xml', 'shape': 2,
        'recipe': [[expr]], 'grid': [[(1, '1', 'General')]], 'source': SimpleNamespace(path=path),
    }
