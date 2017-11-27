import sys
from PyQt5 import QtWidgets, uic
import xlsbatch
import logging
import configparser
import signal
import os.path
import aligntool

Ui_MainWindow, QtBaseClass = uic.loadUiType(
    os.path.join(
        os.path.dirname(os.path.realpath(__file__)), "simplegui.ui")
)
window = None


class AlignToolGui(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        try:
            with open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "VERSION.TXT")) as f:
                versionid = f.readline(10)[:-1]
                self.setWindowTitle("Aligntool GUI version %s" % versionid)
        except IOError:
            pass
        self.browse_xlsx_btn.clicked.connect(self.on_browse_xlsx)
        self.browse_wav_btn.clicked.connect(self.on_browse_wavdir)
        self.quit_btn.clicked.connect(self.on_quit)
        self.export_tg_btn.clicked.connect(self.on_export_tg_btn)
        self.import_tg_btn.clicked.connect(self.on_import_tg_btn)
        self.find_wav_btn.clicked.connect(self.on_find_wav_btn)
        self.segmentBeeps_btn.clicked.connect(self.on_segment_beeps_btn)
        self.addTier_btn.clicked.connect(self.on_add_tier_btn)
        self.segmentSpeechBtn.clicked.connect(self.on_segment_speech_btn)
        self.alignMaus_btn.clicked.connect(self.on_align_maus_btn)
        self.extract_onOffsets_btn.clicked.connect(self.on_extract_on_offsets_btn)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.inipath = os.path.join(os.path.expanduser("~"), ".config", "aligntool.ini")
        cnf = configparser.ConfigParser()
        cnf.read([self.inipath])
        self.xlsxfile_edit.setText(cnf.get('aligntoolgui', 'xlsxfile', fallback=''))
        self.wavdir_edit.setText(cnf.get('aligntoolgui', 'wavdir', fallback=''))
        self.last_xlsx_dir = cnf.get('aligntoolgui', 'lastxlsxdir', fallback='.')
        self.last_wav_dir = cnf.get('aligntoolgui', 'lastwavdir', fallback='.')

    def on_browse_xlsx(self):
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Open file', self.last_xlsx_dir,
                                                            '*.xlsx', '*.xlsx', QtWidgets.QFileDialog.DontConfirmOverwrite)
        if fileName:
            self.xlsxfile_edit.setText(fileName)
            self.last_xlsx_dir = os.path.dirname(fileName)

    def on_browse_wavdir(self):
        wavdir = QtWidgets.QFileDialog.getExistingDirectory(self, 'Directory containing wav files', self.last_wav_dir)
        if wavdir:
            self.wavdir_edit.setText(wavdir)
            self.last_wav_dir = wavdir

    def on_quit(self):
        cnf = configparser.ConfigParser()
        cnf.add_section('aligntoolgui')
        cnf.set('aligntoolgui', 'xlsxfile', self.xlsxfile_edit.text())
        cnf.set('aligntoolgui', 'wavdir', self.wavdir_edit.text())
        cnf.set('aligntoolgui', 'lastxlsxdir', self.last_xlsx_dir)
        cnf.set('aligntoolgui', 'lastwavdir', self.last_wav_dir)
        with open(self.inipath, 'w') as f:
            cnf.write(f)
        self.close()

    def on_find_wav_btn(self):
        runner = xlsbatch.WavImporter()
        self.run_cmd(runner, directory=self.wavdir_edit.text(), xlsxfile=self.xlsxfile_edit.text())

    def on_segment_beeps_btn(self):
        runner = xlsbatch.BatchRunner()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text(), batchcmd="segmentBeeps")

    def on_add_tier_btn(self):
        runner = xlsbatch.BatchRunner()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text(), batchcmd="addTier")

    def on_segment_speech_btn(self):
        runner = xlsbatch.BatchRunner()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text(), batchcmd="segmentSpeech")

    def on_align_maus_btn(self):
        runner = xlsbatch.BatchRunner()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text(), batchcmd="alignMAUS")

    def on_extract_on_offsets_btn(self):
        runner = xlsbatch.OnOffsetExtractor()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text())

    def on_export_tg_btn(self):
        runner = xlsbatch.TextGridBulkExporter()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text())

    def on_import_tg_btn(self):
        runner = xlsbatch.TextGridBulkImporter()
        self.run_cmd(runner, xlsxfile=self.xlsxfile_edit.text())

    def run_cmd(self, task, **kwargs):
        try:
            self.setEnabled(False)
            self.repaint()
            logging.info("running %s %s" % (type(task).__name__, kwargs))
            aligntool.cachedir = os.path.dirname(self.xlsxfile_edit.text())
            task.run(**kwargs)
            logging.info("finished %s %s" % (type(task).__name__, kwargs))
        finally:
            self.setEnabled(True)


def on_app_exit():
    window.on_quit()


def setup():
    signal.signal(signal.SIGINT, signal.SIG_DFL)
    app = QtWidgets.QApplication([])
    global window
    window = AlignToolGui()
    window.show()
    app.aboutToQuit.connect(on_app_exit)
    sys.exit(app.exec_())
