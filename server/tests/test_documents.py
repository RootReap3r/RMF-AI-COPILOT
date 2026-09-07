"""Real document parsing fixtures, generated locally without network calls."""
import pathlib
import subprocess
import tempfile
import unittest
import zipfile

ROOT=pathlib.Path(__file__).resolve().parents[1]
def pdf():
    stream=b'BT /F1 12 Tf 72 720 Td (safe text) Tj ET'
    objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',b'<< /Length '+str(len(stream)).encode()+b' >>\nstream\n'+stream+b'\nendstream']
    data=b'%PDF-1.4\n';offsets=[]
    for n,obj in enumerate(objects,1):
        offsets.append(len(data));data+=f'{n} 0 obj\n'.encode()+obj+b'\nendobj\n'
    xref=len(data);data+=b'xref\n0 6\n0000000000 65535 f \n'
    for offset in offsets:data+=f'{offset:010d} 00000 n \n'.encode()
    return data+f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF'.encode()
class DocumentTests(unittest.TestCase):
    def test_real_formats_and_output_cap(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory=pathlib.Path(tmp)
            (directory/'test.pdf').write_bytes(pdf())
            with zipfile.ZipFile(directory/'test.docx','w') as archive:
                archive.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
                archive.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>safe text</w:t></w:r></w:p></w:body></w:document>')
            for extension in ['.pdf','.docx']:
                proc=subprocess.run(['node',str(ROOT/'parse-worker.js'),str(directory/('test'+extension)),extension],capture_output=True,timeout=15)
                self.assertEqual(proc.returncode,0,proc.stderr)
                self.assertIn(b'safe text',proc.stdout)
            (directory/'large.txt').write_text('x'*(2*1024*1024+1))
            proc=subprocess.run(['node',str(ROOT/'parse-worker.js'),str(directory/'large.txt'),'.txt'],capture_output=True,timeout=15)
            self.assertNotEqual(proc.returncode,0)
if __name__=='__main__':unittest.main()
