import json
import pandas as pd

# Nama file output Excel
OUTPUT_EXCEL_FILENAME = "Hasil_Konversi_Data_LDTI.xlsx"

# ==============================================================================
# ⚠️ PERHATIAN:
# Struktur JSON Anda tidak valid. Data di bawah ini ADALAH data dari file Anda
# yang sudah diperbaiki agar sesuai dengan standar JSON dan bisa di-parse.
# ==============================================================================
CORRECTED_JSON_STRING = """
{
  "sensor jaringan": [
    {
      "id": 1,
      "nama_sensor": "Sensor Suhu 1",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 27.5,
      "timestamp": "2024-04-27T08:00:00Z"
    },
    {
      "id": 2,
      "nama_sensor": "Sensor Kelembaban 1",
      "tipe_sensor": "Kelembaban",
      "status": "aktif",
      "nilai": 60,
      "timestamp": "2024-04-27T08:01:00Z"
    },
    {
      "id": 3,
      "nama_sensor": "Sensor Tekanan 1",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 101.3,
      "timestamp": "2024-04-27T08:02:00Z"
    },
    {
      "id": 4,
      "nama_sensor": "Sensor Suhu 2",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 28.1,
      "timestamp": "2024-04-27T08:03:00Z"
    },
    {
      "id": 5,
      "nama_sensor": "Sensor Kelembaban 2",
      "tipe_sensor": "Kelembaban",
      "status": "nonaktif",
      "nilai": 58,
      "timestamp": "2024-04-27T08:04:00Z"
    },
    {
      "id": 6,
      "nama_sensor": "Sensor Tekanan 2",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 100.8,
      "timestamp": "2024-04-27T08:05:00Z"
    },
    {
      "id": 7,
      "nama_sensor": "Sensor Suhu 3",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 26.9,
      "timestamp": "2024-04-27T08:06:00Z"
    },
    {
      "id": 8,
      "nama_sensor": "Sensor Kelembaban 3",
      "tipe_sensor": "Kelembaban",
      "status": "aktif",
      "nilai": 62,
      "timestamp": "2024-04-27T08:07:00Z"
    },
    {
      "id": 9,
      "nama_sensor": "Sensor Tekanan 3",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 101.0,
      "timestamp": "2024-04-27T08:08:00Z"
    },
    {
      "id": 10,
      "nama_sensor": "Sensor Suhu 4",
      "tipe_sensor": "Suhu",
      "status": "nonaktif",
      "nilai": 27.0,
      "timestamp": "2024-04-27T08:09:00Z"
    },
    {
      "id": 11,
      "nama_sensor": "Sensor Kelembaban 4",
      "tipe_sensor": "Kelembaban",
      "status": "aktif",
      "nilai": 59,
      "timestamp": "2024-04-27T08:10:00Z"
    },
    {
      "id": 12,
      "nama_sensor": "Sensor Tekanan 4",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 101.5,
      "timestamp": "2024-04-27T08:11:00Z"
    },
    {
      "id": 13,
      "nama_sensor": "Sensor Suhu 5",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 27.3,
      "timestamp": "2024-04-27T08:12:00Z"
    },
    {
      "id": 14,
      "nama_sensor": "Sensor Kelembaban 5",
      "tipe_sensor": "Kelembaban",
      "status": "aktif",
      "nilai": 61,
      "timestamp": "2024-04-27T08:13:00Z"
    },
    {
      "id": 15,
      "nama_sensor": "Sensor Tekanan 5",
      "tipe_sensor": "Tekanan",
      "status": "nonaktif",
      "nilai": 100.5,
      "timestamp": "2024-04-27T08:14:00Z"
    },
    {
      "id": 16,
      "nama_sensor": "Sensor Suhu 6",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 28.0,
      "timestamp": "2024-04-27T08:15:00Z"
    },
    {
      "id": 17,
      "nama_sensor": "Sensor Kelembaban 6",
      "tipe_sensor": "Kelembaban",
      "status": "aktif",
      "nilai": 63,
      "timestamp": "2024-04-27T08:16:00Z"
    },
    {
      "id": 18,
      "nama_sensor": "Sensor Tekanan 6",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 101.2,
      "timestamp": "2024-04-27T08:17:00Z"
    },
    {
      "id": 19,
      "nama_sensor": "Sensor Suhu 7",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 27.8,
      "timestamp": "2024-04-27T08:18:00Z"
    },
    {
      "id": 20,
      "nama_sensor": "Sensor Kelembaban 7",
      "tipe_sensor": "Kelembaban",
      "status": "aktif",
      "nilai": 60,
      "timestamp": "2024-04-27T08:19:00Z"
    },
    {
      "id": 21,
      "nama_sensor": "Sensor Tekanan 7",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 101.1,
      "timestamp": "2024-04-27T08:20:00Z"
    },
    {
      "id": 22,
      "nama_sensor": "Sensor Suhu 8",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 27.6,
      "timestamp": "2024-04-27T08:21:00Z"
    },
    {
      "id": 23,
      "nama_sensor": "Sensor Kelembaban 8",
      "tipe_sensor": "Kelembaban",
      "status": "nonaktif",
      "nilai": 59,
      "timestamp": "2024-04-27T08:22:00Z"
    },
    {
      "id": 24,
      "nama_sensor": "Sensor Tekanan 8",
      "tipe_sensor": "Tekanan",
      "status": "aktif",
      "nilai": 100.9,
      "timestamp": "2024-04-27T08:23:00Z"
    },
    {
      "id": 25,
      "nama_sensor": "Sensor Suhu 9",
      "tipe_sensor": "Suhu",
      "status": "aktif",
      "nilai": 27.4,
      "timestamp": "2024-04-27T08:24:00Z"
    }
  ],
  "Laporan shinobi": [
    {
      "id_laporan": 1,
      "id_pengguna": "ANBU_001",
      "tipe_laporan": "teks",
      "isi_laporan": "Musuh terlihat di area utara desa.",
      "koordinat": {"lat": -6.200000, "lon": 106.816666},
      "timestamp": "2024-04-27T08:00:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 2,
      "id_pengguna": "Patroli_005",
      "tipe_laporan": "gambar",
      "isi_laporan": "foto_musuh_0423.jpg",
      "koordinat": {"lat": -6.201000, "lon": 106.817000},
      "timestamp": "2024-04-27T08:05:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 3,
      "id_pengguna": "ANBU_002",
      "tipe_laporan": "suara",
      "isi_laporan": "rekaman_suara_0423.mp3",
      "koordinat": {"lat": -6.202000, "lon": 106.818000},
      "timestamp": "2024-04-27T08:10:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 4,
      "id_pengguna": "Patroli_007",
      "tipe_laporan": "koordinat",
      "isi_laporan": "Musuh terdeteksi di koordinat ini.",
      "koordinat": {"lat": -6.203000, "lon": 106.819000},
      "timestamp": "2024-04-27T08:15:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 5,
      "id_pengguna": "ANBU_003",
      "tipe_laporan": "teks",
      "isi_laporan": "Kondisi medan sulit, waspada terhadap jebakan.",
      "koordinat": {"lat": -6.204000, "lon": 106.820000},
      "timestamp": "2024-04-27T08:20:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 6,
      "id_pengguna": "Patroli_002",
      "tipe_laporan": "gambar",
      "isi_laporan": "jebakan_terlihat_0423.jpg",
      "koordinat": {"lat": -6.205000, "lon": 106.821000},
      "timestamp": "2024-04-27T08:25:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 7,
      "id_pengguna": "ANBU_004",
      "tipe_laporan": "suara",
      "isi_laporan": "rekaman_gerakan_musuh_0423.mp3",
      "koordinat": {"lat": -6.206000, "lon": 106.822000},
      "timestamp": "2024-04-27T08:30:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 8,
      "id_pengguna": "Patroli_009",
      "tipe_laporan": "koordinat",
      "isi_laporan": "Posisi musuh terakhir diketahui.",
      "koordinat": {"lat": -6.207000, "lon": 106.823000},
      "timestamp": "2024-04-27T08:35:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 9,
      "id_pengguna": "ANBU_005",
      "tipe_laporan": "teks",
      "isi_laporan": "Musuh mundur ke hutan sebelah timur.",
      "koordinat": {"lat": -6.208000, "lon": 106.824000},
      "timestamp": "2024-04-27T08:40:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 10,
      "id_pengguna": "Patroli_003",
      "tipe_laporan": "gambar",
      "isi_laporan": "bekas_jejak_musuh_0423.jpg",
      "koordinat": {"lat": -6.209000, "lon": 106.825000},
      "timestamp": "2024-04-27T08:45:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 11,
      "id_pengguna": "ANBU_006",
      "tipe_laporan": "suara",
      "isi_laporan": "rekaman_teriakan_0423.mp3",
      "koordinat": {"lat": -6.210000, "lon": 106.826000},
      "timestamp": "2024-04-27T08:50:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 12,
      "id_pengguna": "Patroli_010",
      "tipe_laporan": "koordinat",
      "isi_laporan": "Musuh berkumpul di area ini.",
      "koordinat": {"lat": -6.211000, "lon": 106.827000},
      "timestamp": "2024-04-27T08:55:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 13,
      "id_pengguna": "ANBU_007",
      "tipe_laporan": "teks",
      "isi_laporan": "Perlu penguatan di sisi barat.",
      "koordinat": {"lat": -6.212000, "lon": 106.828000},
      "timestamp": "2024-04-27T09:00:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 14,
      "id_pengguna": "Patroli_004",
      "tipe_laporan": "gambar",
      "isi_laporan": "pos_patroli_0423.jpg",
      "koordinat": {"lat": -6.213000, "lon": 106.829000},
      "timestamp": "2024-04-27T09:05:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 15,
      "id_pengguna": "ANBU_008",
      "tipe_laporan": "suara",
      "isi_laporan": "rekaman_langkah_kaki_0423.mp3",
      "koordinat": {"lat": -6.214000, "lon": 106.830000},
      "timestamp": "2024-04-27T09:10:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 16,
      "id_pengguna": "Patroli_006",
      "tipe_laporan": "koordinat",
      "isi_laporan": "Musuh memasuki area terlarang.",
      "koordinat": {"lat": -6.215000, "lon": 106.831000},
      "timestamp": "2024-04-27T09:15:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 17,
      "id_pengguna": "ANBU_009",
      "tipe_laporan": "teks",
      "isi_laporan": "Kondisi cuaca buruk, operasi dilanjutkan dengan hati-hati.",
      "koordinat": {"lat": -6.216000, "lon": 106.832000},
      "timestamp": "2024-04-27T09:20:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 18,
      "id_pengguna": "Patroli_008",
      "tipe_laporan": "gambar",
      "isi_laporan": "area_terlarang_0423.jpg",
      "koordinat": {"lat": -6.217000, "lon": 106.833000},
      "timestamp": "2024-04-27T09:25:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 19,
      "id_pengguna": "ANBU_010",
      "tipe_laporan": "suara",
      "isi_laporan": "rekaman_pertempuran_0423.mp3",
      "koordinat": {"lat": -6.218000, "lon": 106.834000},
      "timestamp": "2024-04-27T09:30:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 20,
      "id_pengguna": "Patroli_001",
      "tipe_laporan": "koordinat",
      "isi_laporan": "Musuh terakhir terlihat di sini.",
      "koordinat": {"lat": -6.219000, "lon": 106.835000},
      "timestamp": "2024-04-27T09:35:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 21,
      "id_pengguna": "ANBU_011",
      "tipe_laporan": "teks",
      "isi_laporan": "Perlu evakuasi segera di titik ini.",
      "koordinat": {"lat": -6.220000, "lon": 106.836000},
      "timestamp": "2024-04-27T09:40:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 22,
      "id_pengguna": "Patroli_011",
      "tipe_laporan": "gambar",
      "isi_laporan": "evakuasi_0423.jpg",
      "koordinat": {"lat": -6.221000, "lon": 106.837000},
      "timestamp": "2024-04-27T09:45:00Z",
      "status": "menunggu_verifikasi"
    },
    {
      "id_laporan": 23,
      "id_pengguna": "ANBU_012",
      "tipe_laporan": "suara",
      "isi_laporan": "rekaman_pesan_darurat_0423.mp3",
      "koordinat": {"lat": -6.222000, "lon": 106.838000},
      "timestamp": "2024-04-27T09:50:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 24,
      "id_pengguna": "Patroli_012",
      "tipe_laporan": "koordinat",
      "isi_laporan": "Posisi pasukan musuh terbaru.",
      "koordinat": {"lat": -6.223000, "lon": 106.839000},
      "timestamp": "2024-04-27T09:55:00Z",
      "status": "terverifikasi"
    },
    {
      "id_laporan": 25,
      "id_pengguna": "ANBU_013",
      "tipe_laporan": "teks",
      "isi_laporan": "Situasi terkendali, lanjutkan patroli.",
      "koordinat": {"lat": -6.224000, "lon": 106.840000},
      "timestamp": "2024-04-27T10:00:00Z",
      "status": "terverifikasi"
    }
  ],
  "laporan warga": [
    {
      "id_laporan": 1,
      "id_warga": "Warga_001",
      "koordinat": {"lat": -6.200100, "lon": 106.816700},
      "waktu_kejadian": "2024-04-27T08:00:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 2,
      "id_warga": "Warga_002",
      "koordinat": {"lat": -6.200200, "lon": 106.816800},
      "waktu_kejadian": "2024-04-27T08:05:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 3,
      "id_warga": "Warga_003",
      "koordinat": {"lat": -6.200300, "lon": 106.816900},
      "waktu_kejadian": "2024-04-27T08:10:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 4,
      "id_warga": "Warga_004",
      "koordinat": {"lat": -6.200400, "lon": 106.817000},
      "waktu_kejadian": "2024-04-27T08:15:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 5,
      "id_warga": "Warga_005",
      "koordinat": {"lat": -6.200500, "lon": 106.817100},
      "waktu_kejadian": "2024-04-27T08:20:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 6,
      "id_warga": "Warga_006",
      "koordinat": {"lat": -6.200600, "lon": 106.817200},
      "waktu_kejadian": "2024-04-27T08:25:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 7,
      "id_warga": "Warga_007",
      "koordinat": {"lat": -6.200700, "lon": 106.817300},
      "waktu_kejadian": "2024-04-27T08:30:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 8,
      "id_warga": "Warga_008",
      "koordinat": {"lat": -6.200800, "lon": 106.817400},
      "waktu_kejadian": "2024-04-27T08:35:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 9,
      "id_warga": "Warga_009",
      "koordinat": {"lat": -6.200900, "lon": 106.817500},
      "waktu_kejadian": "2024-04-27T08:40:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 10,
      "id_warga": "Warga_010",
      "koordinat": {"lat": -6.201000, "lon": 106.817600},
      "waktu_kejadian": "2024-04-27T08:45:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 11,
      "id_warga": "Warga_011",
      "koordinat": {"lat": -6.201100, "lon": 106.817700},
      "waktu_kejadian": "2024-04-27T08:50:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 12,
      "id_warga": "Warga_012",
      "koordinat": {"lat": -6.201200, "lon": 106.817800},
      "waktu_kejadian": "2024-04-27T08:55:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 13,
      "id_warga": "Warga_013",
      "koordinat": {"lat": -6.201300, "lon": 106.817900},
      "waktu_kejadian": "2024-04-27T09:00:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 14,
      "id_warga": "Warga_014",
      "koordinat": {"lat": -6.201400, "lon": 106.818000},
      "waktu_kejadian": "2024-04-27T09:05:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 15,
      "id_warga": "Warga_015",
      "koordinat": {"lat": -6.201500, "lon": 106.818100},
      "waktu_kejadian": "2024-04-27T09:10:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 16,
      "id_warga": "Warga_016",
      "koordinat": {"lat": -6.201600, "lon": 106.818200},
      "waktu_kejadian": "2024-04-27T09:15:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 17,
      "id_warga": "Warga_017",
      "koordinat": {"lat": -6.201700, "lon": 106.818300},
      "waktu_kejadian": "2024-04-27T09:20:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 18,
      "id_warga": "Warga_018",
      "koordinat": {"lat": -6.201800, "lon": 106.818400},
      "waktu_kejadian": "2024-04-27T09:25:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 19,
      "id_warga": "Warga_019",
      "koordinat": {"lat": -6.201900, "lon": 106.818500},
      "waktu_kejadian": "2024-04-27T09:30:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 20,
      "id_warga": "Warga_020",
      "koordinat": {"lat": -6.202000, "lon": 106.818600},
      "waktu_kejadian": "2024-04-27T09:35:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 21,
      "id_warga": "Warga_021",
      "koordinat": {"lat": -6.202100, "lon": 106.818700},
      "waktu_kejadian": "2024-04-27T09:40:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 22,
      "id_warga": "Warga_022",
      "koordinat": {"lat": -6.202200, "lon": 106.818800},
      "waktu_kejadian": "2024-04-27T09:45:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 23,
      "id_warga": "Warga_023",
      "koordinat": {"lat": -6.202300, "lon": 106.818900},
      "waktu_kejadian": "2024-04-27T09:50:00Z",
      "status": "terkirim"
    },
    {
      "id_laporan": 24,
      "id_warga": "Warga_024",
      "koordinat": {"lat": -6.202400, "lon": 106.819000},
      "waktu_kejadian": "2024-04-27T09:55:00Z",
      "status": "ditangani"
    },
    {
      "id_laporan": 25,
      "id_warga": "Warga_025",
      "koordinat": {"lat": -6.202500, "lon": 106.819100},
      "waktu_kejadian": "2024-04-27T10:00:00Z",
      "status": "terkirim"
    }
  ],
  "data historis": [
    {
      "id_serangan": 1,
      "nama_serangan": "Invasi Pain ke Konoha",
      "pelaku": "Pain",
      "lokasi": "Konoha",
      "tanggal": "2006-09-15",
      "deskripsi": "Serangan besar-besaran oleh Pain menggunakan teknik Deva Path."
    },
    {
      "id_serangan": 2,
      "nama_serangan": "Serangan Otsutsuki di Konoha",
      "pelaku": "Otsutsuki Kinshiki",
      "lokasi": "Konoha",
      "tanggal": "2023-03-10",
      "deskripsi": "Invasi oleh Kinshiki Otsutsuki dengan kekuatan luar biasa."
    },
    {
      "id_serangan": 3,
      "nama_serangan": "Serangan Pain ke Desa Suna",
      "pelaku": "Pain",
      "lokasi": "Suna",
      "tanggal": "2006-10-01",
      "deskripsi": "Serangan mendadak yang menyebabkan kerusakan besar."
    },
    {
      "id_serangan": 4,
      "nama_serangan": "Invasi Otsutsuki di Desa Kiri",
      "pelaku": "Otsutsuki Momoshiki",
      "lokasi": "Kiri",
      "tanggal": "2023-04-05",
      "deskripsi": "Momoshiki menyerang dengan teknik chakra yang kuat."
    },
    {
      "id_serangan": 5,
      "nama_serangan": "Serangan Pain ke Desa Kumo",
      "pelaku": "Pain",
      "lokasi": "Kumo",
      "tanggal": "2006-11-20",
      "deskripsi": "Serangan terkoordinasi yang menargetkan pusat desa."
    },
    {
      "id_serangan": 6,
      "nama_serangan": "Serangan Otsutsuki di Lembah",
      "pelaku": "Otsutsuki Kinshiki",
      "lokasi": "Lembah",
      "tanggal": "2023-05-12",
      "deskripsi": "Invasi dengan serangan cepat dan destruktif."
    },
    {
      "id_serangan": 7,
      "nama_serangan": "Serangan Pain ke Desa Iwa",
      "pelaku": "Pain",
      "lokasi": "Iwa",
      "tanggal": "2006-12-01",
      "deskripsi": "Serangan yang menyebabkan kehancuran infrastruktur."
    },
    {
      "id_serangan": 8,
      "nama_serangan": "Invasi Otsutsuki di Konoha",
      "pelaku": "Otsutsuki Urashiki",
      "lokasi": "Konoha",
      "tanggal": "2023-06-18",
      "deskripsi": "Serangan mendadak dengan kemampuan teleportasi."
    },
    {
      "id_serangan": 9,
      "nama_serangan": "Serangan Pain ke Desa Yuki",
      "pelaku": "Pain",
      "lokasi": "Yuki",
      "tanggal": "2007-01-10",
      "deskripsi": "Serangan yang menimbulkan korban jiwa besar."
    },
    {
      "id_serangan": 10,
      "nama_serangan": "Invasi Otsutsuki di Desa Suna",
      "pelaku": "Otsutsuki Kinshiki",
      "lokasi": "Suna",
      "tanggal": "2023-07-22",
      "deskripsi": "Serangan dengan teknik chakra yang belum dikenal."
    },
    {
      "id_serangan": 11,
      "nama_serangan": "Serangan Pain ke Desa Konoha",
      "pelaku": "Pain",
      "lokasi": "Konoha",
      "tanggal": "2006-09-16",
      "deskripsi": "Serangan lanjutan dengan penggunaan teknik Shinra Tensei."
    },
    {
      "id_serangan": 12,
      "nama_serangan": "Invasi Otsutsuki di Desa Kumo",
      "pelaku": "Otsutsuki Momoshiki",
      "lokasi": "Kumo",
      "tanggal": "2023-08-30",
      "deskripsi": "Serangan dengan kekuatan fisik dan chakra yang besar."
    },
    {
      "id_serangan": 13,
      "nama_serangan": "Serangan Pain ke Desa Kiri",
      "pelaku": "Pain",
      "lokasi": "Kiri",
      "tanggal": "2006-10-15",
      "deskripsi": "Serangan yang menimbulkan kerusakan parah."
    },
    {
      "id_serangan": 14,
      "nama_serangan": "Invasi Otsutsuki di Lembah",
      "pelaku": "Otsutsuki Urashiki",
      "lokasi": "Lembah",
      "tanggal": "2023-09-10",
      "deskripsi": "Serangan dengan kemampuan manipulasi waktu."
    },
    {
      "id_serangan": 15,
      "nama_serangan": "Serangan Pain ke Desa Yuki",
      "pelaku": "Pain",
      "lokasi": "Yuki",
      "tanggal": "2007-01-20",
      "deskripsi": "Serangan yang menyebabkan kehancuran besar."
    },
    {
      "id_serangan": 16,
      "nama_serangan": "Invasi Otsutsuki di Desa Iwa",
      "pelaku": "Otsutsuki Kinshiki",
      "lokasi": "Iwa",
      "tanggal": "2023-10-05",
      "deskripsi": "Serangan dengan teknik chakra yang kuat dan destruktif."
    },
    {
      "id_serangan": 17,
      "nama_serangan": "Serangan Pain ke Desa Suna",
      "pelaku": "Pain",
      "lokasi": "Suna",
      "tanggal": "2006-10-05",
      "deskripsi": "Serangan mendadak yang menimbulkan korban jiwa."
    },
    {
      "id_serangan": 18,
      "nama_serangan": "Invasi Otsutsuki di Konoha",
      "pelaku": "Otsutsuki Momoshiki",
      "lokasi": "Konoha",
      "tanggal": "2023-11-12",
      "deskripsi": "Serangan dengan teknik chakra yang belum dikenal."
    },
    {
      "id_serangan": 19,
      "nama_serangan": "Serangan Pain ke Desa Kumo",
      "pelaku": "Pain",
      "lokasi": "Kumo",
      "tanggal": "2006-11-25",
      "deskripsi": "Serangan terkoordinasi yang menargetkan pusat desa."
    },
    {
      "id_serangan": 20,
      "nama_serangan": "Invasi Otsutsuki di Desa Kiri",
      "pelaku": "Otsutsuki Urashiki",
      "lokasi": "Kiri",
      "tanggal": "2023-12-01",
      "deskripsi": "Serangan dengan kemampuan teleportasi dan manipulasi chakra."
    },
    {
      "id_serangan": 21,
      "nama_serangan": "Serangan Pain ke Desa Iwa",
      "pelaku": "Pain",
      "lokasi": "Iwa",
      "tanggal": "2006-12-10",
      "deskripsi": "Serangan yang menyebabkan kehancuran infrastruktur."
    },
    {
      "id_serangan": 22,
      "nama_serangan": "Invasi Otsutsuki di Lembah",
      "pelaku": "Otsutsuki Kinshiki",
      "lokasi": "Lembah",
      "tanggal": "2024-01-15",
      "deskripsi": "Invasi dengan serangan cepat dan destruktif."
    },
    {
      "id_serangan": 23,
      "nama_serangan": "Serangan Pain ke Desa Yuki",
      "pelaku": "Pain",
      "lokasi": "Yuki",
      "tanggal": "2007-01-30",
      "deskripsi": "Serangan yang menimbulkan korban jiwa besar."
    },
    {
      "id_serangan": 24,
      "nama_serangan": "Invasi Otsutsuki di Desa Suna",
      "pelaku": "Otsutsuki Momoshiki",
      "lokasi": "Suna",
      "tanggal": "2024-02-20",
      "deskripsi": "Serangan dengan teknik chakra yang belum dikenal."
    },
    {
      "id_serangan": 25,
      "nama_serangan": "Serangan Pain ke Konoha",
      "pelaku": "Pain",
      "lokasi": "Konoha",
      "tanggal": "2006-09-20",
      "deskripsi": "Serangan lanjutan dengan penggunaan teknik Shinra Tensei."
    }
  ]
}
"""

def flatten_data(data, prefix=''):
    """Fungsi pembantu untuk meratakan (flatten) dictionary bersarang (seperti 'koordinat')"""
    flattened = {}
    for key, value in data.items():
        if isinstance(value, dict):
            # Rekursif untuk dictionary bersarang, tambahkan prefix
            nested_data = flatten_data(value, f'{prefix}{key}_')
            flattened.update(nested_data)
        else:
            flattened[f'{prefix}{key}'] = value
    return flattened

def convert_json_to_excel(json_data_string, output_filename):
    """
    Mengkonversi string JSON multi-bagian yang sudah diperbaiki menjadi file Excel.
    """
    try:
        # Load data dari string JSON yang sudah diperbaiki
        data = json.loads(json_data_string)
    except json.JSONDecodeError as e:
        print(f"Error fatal saat mengurai JSON. Struktur yang diperbaiki mungkin masih salah: {e}")
        return

    # Membuat objek ExcelWriter untuk menulis ke file .xlsx
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        print(f"Mulai menulis ke {output_filename}...")
        
        # Iterasi melalui setiap bagian data (setiap sheet)
        for sheet_name, records in data.items():
            if not isinstance(records, list):
                print(f"Melewati '{sheet_name}' karena bukan array data.")
                continue

            # 1. Meratakan data bersarang (koordinat menjadi koordinat_lat & koordinat_lon)
            flattened_records = [flatten_data(record) for record in records]
            
            # 2. Membuat DataFrame Pandas
            df = pd.DataFrame(flattened_records)
            
            # 3. Menulis ke sheet Excel
            clean_sheet_name = sheet_name.replace(" ", "_").strip()
            # Pembatasan nama sheet maksimal 31 karakter
            df.to_excel(writer, sheet_name=clean_sheet_name[:31], index=False)
            print(f"  ✅ Sheet '{clean_sheet_name[:31]}' ({len(df)} baris) berhasil ditulis.")

    print("\nKonversi JSON ke Excel selesai! File Excel Anda siap. 🎉")

# Panggil fungsi untuk konversi
convert_json_to_excel(CORRECTED_JSON_STRING, OUTPUT_EXCEL_FILENAME)