from fbs_runtime.application_context.PyQt5 import ApplicationContext

from PyQt5.QtWidgets import (
    QWidget,
    QLabel,
    QGridLayout,
    QLineEdit,
    QPushButton,
    QProgressBar,
)

from pathlib import Path
from matplotlib import pyplot
import seaborn
import os

# import probscale


class Wykresy(QWidget):
    def __init__(self, parent=None):
        super(Wykresy, self).__init__(parent)
        self.interfejs()

    def interfejs(self):

        self.napis = QLabel("Podaj tydzień:", self)
        self.tydzienBox = QLineEdit(self)
        self.przycisk = QPushButton("Stwórz wykresy", self)
        self.przycisk.clicked.connect(self.poKliknieciu)

        self.status = QProgressBar(self)
        self.status.setMaximum(51)

        self.uklad = QGridLayout()
        self.uklad.addWidget(self.napis, 0, 0)
        self.uklad.addWidget(self.tydzienBox, 0, 1)
        self.uklad.addWidget(self.przycisk, 0, 2)
        self.uklad.addWidget(self.status, 0, 3)
        self.setLayout(self.uklad)

        self.resize(500, 100)
        self.setWindowTitle("Wykresy dla E470_100mm")
        self.show()

    def poKliknieciu(self):
        from pandas import read_excel

        self.tests_limits = (
            "100A 30DG Lag KVAR Test (R)",
            1.85,
            "100A Zero KVAR Test (R)",
            1.85,
            "20A Zero KVAR Test (R)",
            1.85,
            "2A Zero KVAR Test (R)",
            1.85,
            "100A 0.8Lead wh Test (R)",
            0.85,
            "100A Lag Wh Test (R)",
            0.85,
            "100A Unity Wh Test (R)",
            0.85,
            "20A Lag Wh Test (R)",
            0.85,
            "20A Unity Wh Test (R)",
            0.85,
            "20A 0.8Lead wh Test (R)",
            0.85,
            "2A 0.8Lead wh Test (R)",
            0.85,
            "2A Lag Wh Test (R)",
            0.85,
            "2A Unity Wh Test (R)",
            0.85,
            "1A Reverse Wh Test (R)",
            1.35,
            "1A Unity Wh Test (R)",
            1.35,
            "80mA Unity Wh Test (R)",
            20,
            "80mA Unity Wh Test Rev (R)",
            20,
        )
        self.jakiStatus = 0

        tydzien = self.tydzienBox.text()

        # zaladowanie pliku do pandas
        sciezka = Path(
            "//kwiapp10/L G Certification Database/04 HASS/Results/Approvals/"
            + str(tydzien)
            + " 5424A - E470 SMETS2(100mm) 20-100A_Approvals.xlsx"
        )
        self.df = read_excel(sciezka)
        self.df = self.df.reset_index()

        # Setting chart style
        seaborn.set()
        seaborn.set_style("whitegrid")
        seaborn.set_context("paper")

        # Chart size
        self.fig, self.ax = pyplot.subplots(figsize=(8, 6))

        # Utworzenie sciezki
        self.sciezka_do_wykresow = Path(
            "//kwiapp10/L G Certification Database/04 HASS/Wykresy/"
            + str(tydzien)
            + "/"
        )
        if not os.path.exists(str(self.sciezka_do_wykresow)):
            os.makedirs(str(self.sciezka_do_wykresow))

        self.utworzenieTimeplot(tydzien)
        self.utworzenieHistogram(tydzien)
        self.koncowka()

    def koncowka(self):
        self.napis1 = QLabel("Pliki zapisano w:" + str(self.sciezka_do_wykresow), self)
        self.uklad.addWidget(self.napis1, 2, 0)

    def utworzenieTimeplot(self, tydzien):
        i = 0
        while i < len(self.tests_limits):

            self.ax.set(ylabel="Error value", title=self.tests_limits[i])

            self.ax.axes.axhline(
                self.tests_limits[i + 1], linestyle="-.", linewidth=2, color="#CB333B"
            )

            self.ax.axes.axhline(
                -(self.tests_limits[i + 1]),
                linestyle="-.",
                linewidth=2,
                color="#CB333B",
            )

            self.ax.axes.axhline(0, linestyle="-", linewidth=1, color="#008651")

            self.ax.tick_params(
                axis="x", which="both", bottom=False, top=False, labelbottom=False
            )

            self.ax.plot(
                self.df["index"],
                self.df[self.tests_limits[i]],
                linewidth=3,
                alpha=0.75,
                color="#15BEF0",
            )

            self.ax.get_figure().savefig(
                str(self.sciezka_do_wykresow)
                + "/"
                + str(tydzien)
                + " Timeplot "
                + self.tests_limits[i]
                + ".png"
            )
            pyplot.cla()
            i += 2
            self.jakiStatus += 1
            self.status.setValue(self.jakiStatus)

    def utworzenieHistogram(self, tydzien):
        from scipy.stats import norm

        hist_options = dict(color="#15BEF0", alpha=0.5, linewidth=1)

        fit_options = dict(color="#CB333B", alpha=1, linewidth=2)

        i = 0
        while i < len(self.tests_limits):

            self.ax.set(xlabel="", ylabel="Frequency", title=self.tests_limits[i])

            self.ax = seaborn.distplot(
                self.df[self.tests_limits[i]],
                fit=norm,
                kde=False,
                axlabel="",
                hist_kws=hist_options,
                fit_kws=fit_options,
            )

            self.ax.get_figure().savefig(
                str(self.sciezka_do_wykresow)
                + "/"
                + str(tydzien)
                + " Histogram "
                + self.tests_limits[i]
                + ".png"
            )

            pyplot.cla()
            i += 2
            self.jakiStatus += 1
            self.status.setValue(self.jakiStatus)


if __name__ == "__main__":
    import sys

    aplikacja = ApplicationContext()
    okno = Wykresy()
    zamkniecie = aplikacja.app.exec_()
    sys.exit(zamkniecie)
