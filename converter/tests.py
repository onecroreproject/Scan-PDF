import os
from io import BytesIO

from django.test import SimpleTestCase

from .utils import format_download_name, save_uploaded_file


class FormatDownloadNameTests(SimpleTestCase):
    def test_replaces_input_extension_with_converted_extension(self):
        self.assertEqual(
            format_download_name('report.docx', 'C:/temp/scanpdf_outputs/report.pdf'),
            'report.pdf',
        )

    def test_keeps_original_name_when_output_extension_is_missing(self):
        self.assertEqual(
            format_download_name('report.docx', None),
            'report.docx',
        )


class SaveUploadedFileTests(SimpleTestCase):
    def test_save_uploaded_file_reads_from_start_after_seek(self):
        class StubUpload:
            def __init__(self, payload):
                self.name = 'demo.docx'
                self._payload = payload
                self._position = 0

            def seek(self, position):
                self._position = position

            def chunks(self):
                yield self._payload[self._position:]

        upload = StubUpload(b'PK\x03\x04-docx')
        upload.seek(0)
        upload.chunks()

        saved_path = save_uploaded_file(upload)

        with open(saved_path, 'rb') as fh:
            self.assertEqual(fh.read(), b'PK\x03\x04-docx')

        os.remove(saved_path)
