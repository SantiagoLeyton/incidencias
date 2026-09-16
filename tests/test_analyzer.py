import os, tempfile, unittest
from openpyxl import Workbook
from analyzer import analyze
from date_parser import parse_date

class AnalyzerTests(unittest.TestCase):
    def test_dates(self):
        self.assertEqual(parse_date("FEEB.14/26F").status, "RECUPERABLE")
        self.assertEqual(parse_date("FEB.30/26F").status, "INVALIDA")
        self.assertEqual(parse_date("???").status, "AMBIGUA")
    def test_positional_context_and_period(self):
        wb=Workbook(); ws=wb.active; ws.title="T1"
        ws.append(["RECIBOS","FECHA","RECIBIDO","RECIBOS FIDUBOGOTA","RECIBO FIDUCIA SEP/2024"])
        ws.append(["C1;C2","ENE.1/24;ABR.28/24F","=100+200","F1","100"])
        ws.append(["CESIONES"])
        ws.append(["RECIBOS","FECHA","RECIBIDO","RECIBOS FIDUBOGOTA","RECIBO FIDUCIA FEB/2026"])
        ws.append(["C3","FEB.30/26;FEEB.14/26F","5","F2","100"])
        f=tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False); f.close(); wb.save(f.name)
        try:
            result=analyze(f.name); kinds={x["kind"] for x in result["incidents"]}
            self.assertIn("HIST_INVALID_DATE_HEADER", kinds)
            self.assertTrue(any(x["context"]=="CESIÓN" for x in result["incidents"]))
        finally: os.unlink(f.name)
    def test_multiple_period_evidences_are_one_level_two_incident(self):
        wb=Workbook(); ws=wb.active
        ws.append(["RECIBOS","FECHA","RECIBIDO","RECIBOS FIDUBOGOTA","RECIBO FIDUCIA SEP/2024"])
        ws.append(["","ABR.1/24F-MAY.1/24F","","F1-F2","=100+200"])
        f=tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False); f.close(); wb.save(f.name)
        try:
            result=analyze(f.name)
            periods=[x for x in result["incidents"] if x["kind"]=="HIST_PAYMENT_VALUE_COUNT_MISMATCH"]
            self.assertEqual(len(periods), 1)
            self.assertEqual(len(periods[0]["evidence_items"]), 2)
            self.assertIn("2", periods[0]["info"][0])
        finally: os.unlink(f.name)
if __name__ == "__main__": unittest.main()
