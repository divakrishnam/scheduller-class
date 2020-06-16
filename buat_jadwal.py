import math
import random as rnd
import prettytable
import winsound
import os
import ssl
import smtplib
import config
import re
import pandas as pd
import threading

from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class Utils:
    regex = re.compile('[@_!#$%^&*()<>?/\|}{~:],;')

    def getRuangan(self):
        ruangans = [['223', 40, 'UMM'],
                    ['224', 40, 'UMM'],
                    ['225', 40, 'UMM'],
                    ['104', 40, 'UMM'],
                    ['105', 40, 'UMM'],
                    ['201', 40, 'UMM'],
                    ['202', 40, 'UMM'],
                    ['203', 40, 'UMM'],
                    ['204', 40, 'UMM'],
                    ['205', 40, 'UMM'],
                    ['206', 40, 'UMM'],
                    ['226', 40, 'UMM'],
                    ['211', 40, 'UMM'],
                    ['212', 40, 'UMM'],
                    ['213', 40, 'UMM'],
                    ['214', 40, 'UMM'],
                    ['108', 40, 'UMM'],
                    ['216', 40, 'UMM'],
                    ['109', 40, 'UMM'],
                    ['301', 40, 'UMM'],
                    ['302', 40, 'UMM'],
                    ['303', 40, 'UMM'],
                    ['304', 40, 'UMM'],
                    ['305', 40, 'UMM'],
                    ['306', 40, 'UMM'],
                    ['307', 40, 'BHS'],
                    ['308', 40, 'BHS'],
                    ['309', 35, 'KOM2'],
                    ['310', 35, 'KOM2'],
                    ['312', 35, 'KOM2'],
                    ['313', 35, 'KOM2'],
                    ['314', 35, 'KOM1'],
                    ['315', 35, 'KOM1'],
                    ['316', 35, 'KOM2'],
                    ['207', 40, 'UMM'],
                    ['107', 40, 'KOM2'],
                    ['110', 40, 'UMM'],
                    ['111', 40, 'UMM'],
                    ['112', 40, 'UMM'],
                    ['114', 40, 'UMM'],
                    ['115', 40, 'UMM'],
                    ['116', 40, 'UMM'],
                    ['117', 40, 'UMM'],
                    ['208', 40, 'UMM'],
                    ['209', 40, 'UMM'],
                    ['217', 40, 'UMM'],
                    ['218', 40, 'UMM'],
                    ['220', 40, 'UMM'],
                    ['221', 40, 'UMM'],
                    ['222', 40, 'UMM'],
                    ['219', 40, 'UMM'],
                    ['LAB 312', 35, 'KOM2'],
                    ['LAB 313', 35, 'KOM2'],
                    ['LAB 314', 35, 'KOM1'],
                    ['LAB 315', 35, 'KOM1'],
                    ['320', 40, 'BHS'],
                    ['321', 40, 'BHS'],
                    ['Lab 309', 35, 'KOM2'],
                    ['Lab 310', 35, 'KOM2'],
                    ['Lab 311', 35, 'KOM2'],
                    ['Lab 316', 35, 'KOM2'],
                    ['Lab 317', 35, 'KOM2'],
                    ['Lab 318', 35, 'KOM3'],
                    ['106', 40, 'UMM'],
                    ['113', 100, 'UMM'],
                    ['101', 40, 'UMM'],
                    ['102', 40, 'UMM'],
                    ['103', 40, 'UMM'],
                    ['210', 40, 'UMM']]
        return ruangans

    def getWaktuMatkuls(self):
        waktus = [['W01', '07:00-08:40', 100],
                 ['W02', '08:40-10:20', 100],
                 ['W03', '10:20-12:00', 100],
                 ['W04', '13:00-14:40', 100],
                 ['W05', '14:40-16:20', 100],
                 ['W06', '16:20-18:00', 100],
                 ['W07', '07:50-09:30', 100],
                 ['W08', '09:30-11:10', 100],
                 ['W09', '13:50-15:30', 100],
                 ['W10', '15:30-17:10', 100],
                 ['W11', '07:00-09:30', 150],
                 ['W12', '09:30-12:00', 150],
                 ['W13', '13:00-15:30', 150],
                 ['W14', '15:30-18:00', 150],
                 ['W15', '07:50-10:20', 150],
                 ['W16', '13:50-16:20', 150],
                 ['W17', '08:40-11:10', 150],
                 ['W18', '14:40-17:10', 150],
                 ['W19', '07:00-10:20', 200],
                 ['W20', '13:00-16:20', 200],
                 ['W21', '07:50-11:10', 200],
                 ['W22', '13:50-17:10', 200],
                 ['W23', '08:40-12:00', 200],
                 ['W24', '14:40-18:00', 200],
                 ['W25', '07:00-11:10', 250],
                 ['W26', '13:00-17:10', 250],
                 ['W27', '07:50-12:00', 250],
                 ['W28', '13:50-18:00', 250],
                 ['W29', '07:00-12:00', 300],
                 ['W30', '13:00-18:00', 300]]
        return waktus

    def getProdis(self):
        prodis = [['13', 'D3 Teknik Informatika', []],
                  ['14', 'D4 Teknik Informatika', []],
                  ['23', 'D3 Manajemen Informatika', []],
                  ['33', 'D3 Akuntansi', []],
                  ['34', 'D4 Akuntansi Keuangan', []],
                  ['43', 'D3 Manajemen Pemasaran', []],
                  ['44', 'D4 Manajemen Perusahaan', []],
                  ['53', 'D3 Logistik Bisnis', []],
                  ['54', 'D4 Logistik Bisnis', []],
                  ['31', 'D1 Akuntansi', []],
                  ['93', 'D3 DikSif', []],
                  ['74', 'D4 E-Commerce Logistics', []]]

        return prodis

    def getMatkuls(self, path):
        df = pd.read_excel(r""+path)
        df = df[df["Perlu Jadwal"] == "Y"]
        kolom = ['SKS', "Perlu Jadwal"]
        df.drop(kolom, axis=1, inplace=True)
        kolom = ["Prodi", "Jumlah Mahasiswa"]
        df[kolom] = df[kolom].applymap(lambda x: '%d' % x)
        kolom = ["Total Jam"]
        df[kolom] = df[kolom].applymap(lambda x: '%d' % (int(x)*50))
        listKode = ["M"+str(x+1).zfill(3) for x in range(0, len(df))]
        df.insert(0, "Kode", listKode, True)
        return df.values.tolist()

    def findItem(self, lists, item):
        return [ind for ind in range(len(lists)) if item in lists[ind]]

    def findListItem(self, lists, item):
        return [lists[ind] for ind in range(len(lists)) if item in lists[ind]]

    def formatWaktu(self, waktu):
        waktu = waktu.replace(" ", "")
        waktus = waktu.split('-')
        awal = datetime.strptime(waktus[0], '%H:%M')
        akhir = datetime.strptime(waktus[1], '%H:%M')
        return [awal, akhir]

    def getDay(self, day):
        dictDay = {
            "1": "senin",
            "2": "selasa",
            "3": "rabu",
            "4": "kamis",
            "5": "jumat",
            "6": "sabtu",
        }
        return dictDay.get(day, "bebas")

    def suaraBeep(self):
        frequency = 2500
        duration = 1000
        for i in range(0, 5):
            winsound.Beep(frequency, duration)

    def sendEmail(self, file, receiver):
        try:
            subject = "Jadwal Baru"
            body = "Ini jadwal barunya boss..)"

            sender_email = config.my_email
            receiver_email = receiver
            password = config.my_pass

            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            message["Bcc"] = receiver_email

            message.attach(MIMEText(body, "plain"))

            with open(file, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)

            part.add_header(
                "Content-Disposition",
                "attachment; filename= %s " % file,
            )

            message.attach(part)

            text = message.as_string()

            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)

            print('File %s berhasil dikirim ke %s' % (file, receiver))

        except FileNotFoundError:
            pass


utils = Utils()


class Data:
    path_abs = os.path.abspath(__file__)
    path_now = os.path.dirname(path_abs)
    file = path_now+"\data_ploting.xlsx"

    HARIS = [["SEN", "Senin"],
             ["SEL", "Selasa"],
             ["RAB", "Rabu"],
             ["KAM", "Kamis"],
             ["JUM", "Jumat"],
             ["SAB", "Sabtu"]]

    RUANGANS = utils.getRuangan()

    WAKTU_MATKULS = utils.getWaktuMatkuls()
    PRODIS = utils.getProdis()
    MATKULS = utils.getMatkuls(file)

    def __init__(self):
        self._ruangans = []
        self._waktuMatkuls = []
        self._haris = []
        self._prodis = []
        self._matkuls = []

        for i in range(0, len(self.RUANGANS)):
            self._ruangans.append(
                Ruangan(self.RUANGANS[i][0],
                        self.RUANGANS[i][1],
                        self.RUANGANS[i][2]))

        for i in range(0, len(self.WAKTU_MATKULS)):
            self._waktuMatkuls.append(
                WaktuMatkul(self.WAKTU_MATKULS[i][0],
                            self.WAKTU_MATKULS[i][1],
                            self.WAKTU_MATKULS[i][2]))

        for i in range(0, len(self.HARIS)):
            self._haris.append(
                Hari(self.HARIS[i][0],
                     self.HARIS[i][1]))

        for i in range(0, len(self.MATKULS)):
            self._matkuls.append(
                Matkul(self.MATKULS[i][0],
                       self.MATKULS[i][1],
                       self.MATKULS[i][2],
                       self.MATKULS[i][3],
                       self.MATKULS[i][4],
                       self.MATKULS[i][5],
                       self.MATKULS[i][6],
                       self.MATKULS[i][7],
                       self.MATKULS[i][8],
                       self.MATKULS[i][9],
                       self.MATKULS[i][10])
            )

        for prodi in self.PRODIS:
            matkuls = utils.findListItem(self.MATKULS, prodi[0])

            if matkuls:
                newMatkuls = []
                for matkul in matkuls:
                    newMatkuls.append(Matkul(
                        matkul[0], matkul[1],
                        matkul[2], matkul[3],
                        matkul[4], matkul[5],
                        matkul[6], matkul[7],
                        matkul[8], matkul[9],
                        matkul[10]
                    ))

                self.PRODIS[self.PRODIS.index(prodi)][2].extend(newMatkuls)

        for i in range(0, len(self.PRODIS)):
            if self.PRODIS[i][2]:
                self._prodis.append(
                    Prodi(self.PRODIS[i][0],
                          self.PRODIS[i][1],
                          self.PRODIS[i][2]))

        self._banyakKelas = 0
        for i in range(0, len(self._prodis)):
            self._banyakKelas += len(self._prodis[i].getMatkul())

    def getRuangans(self):
        return self._ruangans

    def getHaris(self):
        return self._haris

    def getMatkuls(self):
        return self._matkuls

    def getProdis(self):
        return self._prodis

    def getWaktuMatkuls(self):
        return self._waktuMatkuls

    def getBanyakKelas(self):
        return self._banyakKelas


class Ruangan:
    def __init__(self, kode, kapasitas, jenis):
        self._kode = kode
        self._kapasitas = kapasitas
        self._jenis = jenis

    def getKode(self):
        return self._kode

    def getKapasitas(self):
        return self._kapasitas

    def getJenis(self):
        return self._jenis


class Hari:
    def __init__(self, kode, nama):
        self._kode = kode
        self._nama = nama

    def getKode(self):
        return self._kode

    def getNama(self):
        return self._nama


class WaktuMatkul:
    def __init__(self, kode, waktu, menit):
        self._kode = kode
        self._waktu = waktu
        self._menit = menit

    def getKode(self):
        return self._kode

    def getWaktu(self):
        return self._waktu

    def getMenit(self):
        return self._menit


class Prodi:
    def __init__(self, kode, nama, matkul):
        self._kode = kode
        self._nama = nama
        self._matkul = matkul

    def getKode(self):
        return self._kode

    def getNama(self):
        return self._nama

    def getMatkul(self):
        return self._matkul


class Matkul:
    def __init__(self, kode, prodi, kodeMatkul, namaMatkul, menit, kelas, kodeDosen, namaDosen, tipeHari, jenis, maxMahasiswa):
        self._kode = kode
        self._prodi = prodi
        self._kodeMatkul = kodeMatkul
        self._namaMatkul = namaMatkul
        self._menit = menit
        self._kelas = kelas
        self._kodeDosen = kodeDosen
        self._namaDosen = namaDosen
        self._tipeHari = str(tipeHari)
        self._jenis = jenis
        self._maxMahasiswa = maxMahasiswa

    def getKode(self):
        return self._kode

    def getProdi(self):
        return self._prodi

    def getKodeMatkul(self):
        return self._kodeMatkul

    def getNamaMatkul(self):
        return self._namaMatkul

    def getMenit(self):
        return self._menit

    def getKelas(self):
        return self._kelas

    def getTipeHari(self):
        return self._tipeHari

    def getKodeDosen(self):
        return self._kodeDosen

    def getNamaDosen(self):
        return self._namaDosen

    def getJenis(self):
        return self._jenis

    def getMaxMahasiswa(self):
        return self._maxMahasiswa

    def __str__(self):
        return self._kode


class Jadwal:
    def __init__(self, kode, prodi, matkul):
        self._kode = kode
        self._prodi = prodi
        self._matkul = matkul
        self._hari = None
        self._waktuMatkul = None
        self._ruangan = None

    def getKode(self):
        return self._kode

    def getProdi(self):
        return self._prodi

    def getMatkul(self):
        return self._matkul

    def getHari(self):
        return self._hari

    def getWaktuMatkul(self):
        return self._waktuMatkul

    def getRuangan(self):
        return self._ruangan

    def setHari(self, hari):
        self._hari = hari

    def setWaktuMatkul(self, waktuMatkul):
        self._waktuMatkul = waktuMatkul

    def setRuangan(self, ruangan):
        self._ruangan = ruangan

    def __str__(self):
        return self._prodi.getNama() +\
            " | "+self._matkul.getNamaMatkul()+"-"+self._matkul.getKelas() +\
            " | "+self._matkul.getNamaDosen() +\
            " | "+self._ruangan.getKode() + \
            " | "+self._hari.getNama() +\
            " | "+self._waktuMatkul.getWaktu() +\
            " | "+str(self._waktuMatkul.getMenit()) + \
            " | "+str(self._matkul.getMenit())


class Penjadwalan:
    def __init__(self):
        self._data = data
        self._jadwal = []
        self._konflik = 0
        self._fitness = -1
        self._kodeJadwal = 0
        self._fitnessBerubah = True

    def getJadwal(self):
        self._fitnessBerubah = True
        return self._jadwal

    def getKonflik(self):
        return self._konflik

    def getFitness(self):
        if(self._fitnessBerubah == True):
            self._fitness = self.hitungFitness()
            self._fitnessBerubah = False

        return self._fitness

    def inisialisasi(self):
        prodis = self._data.getProdis()

        for i in range(0, len(prodis)):
            matkuls = prodis[i].getMatkul()

            for j in range(0, len(matkuls)):
                jadwalBaru = Jadwal(self._kodeJadwal, prodis[i], matkuls[j])
                self._kodeJadwal += 1
                jadwalBaru.setHari(
                    data.getHaris()[rnd.randrange(0, len(data.getHaris()))])
                jadwalBaru.setWaktuMatkul(
                    data.getWaktuMatkuls()[rnd.randrange(0, len(data.getWaktuMatkuls()))])
                jadwalBaru.setRuangan(
                    data.getRuangans()[rnd.randrange(0, len(data.getRuangans()))])
                self._jadwal.append(jadwalBaru)

        return self

    def hitungFitness(self):
        self._konflik = 0
        jadwal = self.getJadwal()

        message = ''
        for i in range(0, len(jadwal)):
            if (int(jadwal[i].getRuangan().getKapasitas()) < int(jadwal[i].getMatkul().getMaxMahasiswa())):
                self._konflik += 1
                message += "1 "

            if (str(jadwal[i].getRuangan().getJenis()) != str(jadwal[i].getMatkul().getJenis())):
                self._konflik += 1
                message += "2 "

            if (int(jadwal[i].getWaktuMatkul().getMenit()) != int(jadwal[i].getMatkul().getMenit())):
                self._konflik += 1
                message += "3 "
                # print(int(jadwal[i].getWaktuMatkul().getMenit()), int(jadwal[i].getMatkul().getMenit()))

            text = ""
            for x in jadwal[i].getMatkul().getTipeHari():
                # print(x.isalpha())
                if x.isalpha():
                    continue
                else:
                    text += x
            # print(text)

            if ";" in text:
                # print(text+"z")
                days = [str(int(float(x)))
                        for x in text.split(";")]
            elif utils.regex.search(text) == None and text:
                # print(text+"x")
                days = [
                    str(int(float(text)))]
            else:
                # print(text+"y")
                days = [text]

            # print(days)
            listDay = [utils.getDay(day) for day in days]
            # print(days, listDay, jadwal[i].getHari().getNama().lower())

            # print(listDay[0])
            if (jadwal[i].getHari().getNama().lower() not in listDay and listDay[0] != "bebas"):
                self._konflik += 1
                message += "4 "+jadwal[i].getHari().getNama().lower() + str(listDay) +\
                    str(jadwal[i].getHari().getNama().lower()
                        not in listDay) + str(listDay[0] != "bebas")

      # if (jadwal[i].getMatkul().getTipeHari() == "1" and jadwal[i].getHari().getNama().lower() == "selasa"):
      #     self._konflik += 1
      #     message += "4 "

      # if (jadwal[i].getMatkul().getTipeHari() == "2" and jadwal[i].getHari().getNama().lower() != "sabtu"):
      #     self._konflik += 1
      #     message += "5 "

            waktuSatu = utils.formatWaktu(
                jadwal[i].getWaktuMatkul().getWaktu())

            for j in range(0, len(jadwal)):
                if (j >= i):
                    if jadwal[i].getMatkul().getKodeDosen() == jadwal[j].getMatkul().getKodeDosen() and jadwal[i].getHari().getKode() == jadwal[j].getHari().getKode():
                        waktuDua = utils.formatWaktu(
                            jadwal[j].getWaktuMatkul().getWaktu())
                        if(not(waktuSatu[0] <= waktuDua[0] and waktuSatu[1] <= waktuDua[0] and waktuSatu[0] <= waktuDua[1]) and jadwal[i].getKode() != jadwal[j].getKode()):
                            self._konflik += 1
                            message += "6 "
                    else:
                        if jadwal[i].getHari().getKode() == jadwal[j].getHari().getKode() and jadwal[i].getRuangan().getKode() == jadwal[j].getRuangan().getKode():
                            waktuDua = utils.formatWaktu(
                                jadwal[j].getWaktuMatkul().getWaktu())
                            if(not(waktuSatu[0] <= waktuDua[0] and waktuSatu[1] <= waktuDua[0] and waktuSatu[0] <= waktuDua[1]) and jadwal[i].getKode() != jadwal[j].getKode()):
                                self._konflik += 1
                                message += "7 "
            # print(message)
        return 1 / ((1.0*self._konflik + 1))

    def __str__(self):
        nilai = ""
        for i in range(0, len(self._jadwal)-1):
            nilai += str(self._jadwal[i])+", "

        nilai += str(self._jadwal[len(self._jadwal)-1])

        return nilai


class Populasi:
    def __init__(self, ukuran):
        self._ukuran = ukuran
        self._data = data
        self._jadwals = []

        for i in range(0, ukuran):
            self._jadwals.append(Penjadwalan().inisialisasi())

    def getJadwals(self):
        return self._jadwals


class AlgoritmaGenetik:
    def evolve(self, populasi):
        return self._mutasiPopulasi(self._crossoverPopulasi(populasi))

    def _crossoverPopulasi(self, pop):
        crossoverPop = Populasi(0)

        for i in range(NOMOR_JADWAL_ELIT):
            crossoverPop.getJadwals().append(pop.getJadwals()[i])
        i = NOMOR_JADWAL_ELIT

        while i < UKURAN_POPULASI:
            jadwalSatu = self._pilihPertandinganPopulasi(pop).getJadwals()[0]
            jadwalDua = self._pilihPertandinganPopulasi(pop).getJadwals()[0]
            crossoverPop.getJadwals().append(self._crossoverJadwal(jadwalSatu, jadwalDua))
            i += 1
        return crossoverPop

    def _mutasiPopulasi(self, populasi):
        for i in range(NOMOR_JADWAL_ELIT, UKURAN_POPULASI):
            self._mutasiJadwal(populasi.getJadwals()[i])
        return populasi

    def _crossoverJadwal(self, jadwalSatu, jadwalDua):
        crossoverJadwal = Penjadwalan().inisialisasi()

        for i in range(0, len(crossoverJadwal.getJadwal())):
            if(rnd.random() > 0.5):
                crossoverJadwal.getJadwal()[i] = jadwalSatu.getJadwal()[i]
            else:
                crossoverJadwal.getJadwal()[i] = jadwalDua.getJadwal()[i]

        return crossoverJadwal

    def _mutasiJadwal(self, mutasiJadwal):
        jadwal = Penjadwalan().inisialisasi()

        for i in range(0, len(mutasiJadwal.getJadwal())):
            if(TINGKAT_MUTASI > rnd.random()):
                mutasiJadwal.getJadwal()[i] = jadwal.getJadwal()[i]

        return mutasiJadwal

    def _pilihPertandinganPopulasi(self, pop):
        pertandinganPop = Populasi(0)
        i = 0

        while i < UKURAN_PERTANDINGAN_TERPILIH:
            pertandinganPop.getJadwals().append(
                pop.getJadwals()[rnd.randrange(0, UKURAN_POPULASI)])
            i += 1
        pertandinganPop.getJadwals().sort(key=lambda x: x.getFitness(), reverse=True)

        return pertandinganPop


class TampilData:
    def tampilDataTersedia(self):
        print("> Semua Data Yang Tersedia")
        self.tampilHari()
        self.tampilRuangan()
        self.tampilWaktuMatkul()
        self.tampilProdi()
        self.tampilMatkul()

    def tampilHari(self):
        tableHari = prettytable.PrettyTable(
            ['Kode', 'Nama'])
        haris = data.getHaris()

        for i in range(0, len(haris)):
            tableHari.add_row(
                [haris[i].getKode(), haris[i].getNama()])

        print(tableHari)

    def tampilRuangan(self):
        tableRuangans = prettytable.PrettyTable(
            ['Kode', 'Kapasitas', 'Jenis'])
        ruangans = data.getRuangans()

        for i in range(0, len(ruangans)):
            tableRuangans.add_row(
                [str(ruangans[i].getKode()),
                 str(ruangans[i].getKapasitas()),
                 str(ruangans[i].getJenis())])

        print(tableRuangans)

    def tampilWaktuMatkul(self):
        tableWaktuMatkuls = prettytable.PrettyTable(
            ['Kode', 'Waktu', 'Menit'])
        waktuMatkuls = data.getWaktuMatkuls()

        for i in range(0, len(waktuMatkuls)):
            tableWaktuMatkuls.add_row(
                [waktuMatkuls[i].getKode(),
                 waktuMatkuls[i].getWaktu(),
                 str(waktuMatkuls[i].getMenit())])

        print(tableWaktuMatkuls)

    def tampilProdi(self):
        prodis = data.getProdis()
        tabelProdi = prettytable.PrettyTable(['kode', 'nama', 'matkul'])

        for i in range(0, len(prodis)):
            matkuls = prodis.__getitem__(i).getMatkul()
            tempStr = ""
            for j in range(0, len(matkuls)):
                tempStr += matkuls[j].__str__()+"\n "
            tabelProdi.add_row(
                [prodis.__getitem__(i).getNama(), prodis.__getitem__(i).getKode(), tempStr])

        print(tabelProdi)

    def tampilMatkul(self):
        tableMatkuls = prettytable.PrettyTable(
            ['Kode', 'Prodi',
             'Nama Matkul', 'Menit',
             'Kelas', 'Nama Dosen',
             'Jabatan', 'Jenis Ruang',
             'Max Peserta'])

        matkuls = data.getMatkuls()

        for i in range(0, len(matkuls)):
            tableMatkuls.add_row(
                [matkuls[i].getKode(),
                 matkuls[i].getProdi(),
                 matkuls[i].getNamaMatkul(),
                 str(matkuls[i].getMenit()),
                 matkuls[i].getKelas(),
                 matkuls[i].getNamaDosen(),
                 matkuls[i].getTipeHari(),
                 matkuls[i].getJenis(),
                 str(matkuls[i].getMaxMahasiswa())]
            )
        print(tableMatkuls)

    def tampilGenerasiTeratas(self, populasi):
        tabelGenerasis = prettytable.PrettyTable(
            ['No', 'Fitness', 'Konflik', 'Jadwal'])
        jadwals = populasi.getJadwals()

        for i in range(0, 1):
            hasilJadwal = "-"
            # for j in str(jadwals[i]).split(", "):
            #     hasilJadwal += j+"\n"
            tabelGenerasis.add_row(
                [str(i+1),
                 round(jadwals[i].getFitness(), 3),
                 jadwals[i].getKonflik(),
                 hasilJadwal])

        print(tabelGenerasis)

    def tampilGenerasi(self, populasi):
        tabelGenerasis = prettytable.PrettyTable(
            ['No', 'Fitness', 'Konflik', 'Jadwal'])
        jadwals = populasi.getJadwals()

        for i in range(0, len(jadwals)):
            hasilJadwal = "-"
            # for j in str(jadwals[i]).split(", "):
            #     hasilJadwal += j+"\n"
            tabelGenerasis.add_row([str(i+1),
                                    round(jadwals[i].getFitness(), 3),
                                    jadwals[i].getKonflik(),
                                    hasilJadwal])

        print(tabelGenerasis)

    def tampilJadwalKosong(self, dfJadwalKuliah):
        haris = data.getHaris()
        waktuMatkuls = data.getWaktuMatkuls()
        ruangans = data.getRuangans()
        jadwalKosong = []

        for h in range(0, len(haris)):
            hari = haris[h].getNama()
            for w in range(0, len(waktuMatkuls)):
                waktu = waktuMatkuls[w].getWaktu()
                menit = waktuMatkuls[w].getMenit()
                for r in range(0, len(ruangans)):
                    ruangan = ruangans[r].getKode()
                    jenis = ruangans[r].getJenis()
                    jadwalKosong.append([hari, waktu, menit, ruangan, jenis])

        dfJadwalKosong = pd.DataFrame(
            jadwalKosong, columns=["Hari", "Waktu",
                                   "Menit", "Ruang",
                                   "Jenis Ruang"])

        kecuali = ["Hari", "Waktu",
                   "Menit", "Ruang",
                   "Jenis Ruang"]

        koloms = [kolom for kolom in dfJadwalKuliah.columns if kolom not in kecuali]

        dfJadwalKuliah = dfJadwalKuliah.drop(koloms, axis=1)

        def waktuSama(waktu, row):
            if (datetime.strptime(waktu.split("-")[0], '%H:%M') >= datetime.strptime(row.Waktu.split("-")[0], '%H:%M')) and (datetime.strptime(waktu.split("-")[1], '%H:%M') <= datetime.strptime(row.Waktu.split("-")[1], '%H:%M')):
                return waktu

        for row in dfJadwalKuliah.itertuples(index=True, name='Pandas'):
            deleteRows = dfJadwalKosong.loc[(dfJadwalKosong["Hari"] == row.Hari) & (dfJadwalKosong["Ruang"] == row.Ruang) & (
                [waktuSama(waktu, row) for waktu in dfJadwalKosong["Waktu"]])].index
            dfJadwalKosong.drop(deleteRows, inplace=True)

        return dfJadwalKosong

    def tampilJadwalExcel(self, jadwal):
        jadwal = jadwal.getJadwal()

        jadwalBaru = []
        for i in range(0, len(jadwal)):
            jadwalBaru.append([str(i+1), jadwal[i].getProdi().getNama(),
                               jadwal[i].getMatkul().getNamaMatkul(),
                               jadwal[i].getMatkul().getKelas(),
                               jadwal[i].getMatkul().getMenit(),
                               jadwal[i].getMatkul().getNamaDosen(),
                               jadwal[i].getMatkul().getMaxMahasiswa(),
                               jadwal[i].getMatkul().getJenis(),
                               jadwal[i].getHari().getNama(),
                               jadwal[i].getWaktuMatkul().getWaktu(),
                               jadwal[i].getWaktuMatkul().getMenit(),
                               jadwal[i].getRuangan().getKode(),
                               jadwal[i].getRuangan().getJenis()
                               ])

        dfJadwalKuliah = pd.DataFrame(jadwalBaru,
                                      columns=["No.", "Prodi",
                                               "Matkul", "Kelas",
                                               "Menit", "Dosen",
                                               "Max Mahasiswa", "Jenis Ruang",
                                               "Hari", "Waktu", "Menit Asli",
                                               "Ruang", "Jenis Ruang Asli"])

        dfJadwalKosong = self.tampilJadwalKosong(dfJadwalKuliah)

        writer = pd.ExcelWriter("jadwal_kuliah.xlsx", engine='xlsxwriter')

        dfJadwalKuliah.to_excel(
            writer, sheet_name='Jadwal Kuliah', index=False)
        dfJadwalKosong.to_excel(
            writer, sheet_name="Jadwal Kosong", index=False)

        worksheet = writer.sheets['Jadwal Kuliah']
        worksheet.set_column('B:B', 22, None)
        worksheet.set_column('J:J', 14, None)
        writer.save()


data = Data()

tampilData = TampilData()

lenMatkul = len(data.getMatkuls())+2
UKURAN_POPULASI = math.ceil(lenMatkul/2)+lenMatkul
NOMOR_JADWAL_ELIT = 1
UKURAN_PERTANDINGAN_TERPILIH = math.ceil(lenMatkul/2)
TINGKAT_MUTASI = 0.1


def main():
    tampilData.tampilDataTersedia()

    noGenerasi = 0
    print("\n> Generasi "+str(noGenerasi))
    populasi = Populasi(UKURAN_POPULASI)
    populasi.getJadwals().sort(key=lambda x: x.getFitness(), reverse=True)
    tampilData.tampilGenerasiTeratas(populasi)

    algoritmaGenetik = AlgoritmaGenetik()

    while(populasi.getJadwals()[0].getFitness() != 1.0):
        noGenerasi += 1
        print("\n> Generasi "+str(noGenerasi))
        populasi = algoritmaGenetik.evolve(populasi)
        populasi.getJadwals().sort(key=lambda x: x.getFitness(), reverse=True)
        tampilData.tampilGenerasiTeratas(populasi)

    tampilData.tampilJadwalExcel(populasi.getJadwals()[0])

    send = threading.Thread(target=utils.sendEmail, args=("jadwal_kuliah.xlsx","divakrishna55@gmail.com"))
    send.start()

    utils.suaraBeep()
    print("\n")

main()
