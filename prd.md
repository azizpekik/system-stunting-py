⸻

📘 RULES: KONVERSI, KOMPARASI, DAN VALIDASI DATA PERTUMBUHAN ANAK

⸻

A. RULES: PEMBACAAN DATA EXCEL
	1.	Struktur Excel terdiri dari:
	•	Kolom identitas anak (wajib):
	•	NO
	•	NIK
	•	NAMA ANAK
	•	TANGGAL LAHIR
	•	JENIS KELAMIN
	•	Kolom periode bulanan (misal: JANUARI 2024, FEBRUARI 2024, dst.)
	•	Setiap periode berisi 5 sub-kolom (baris kedua Excel):
	•	TGL UKUR
	•	UMUR
	•	BERAT
	•	TINGGI
	•	CARA UKUR
	2.	Data anak dimulai pada baris ke-3.
	3.	Jika sebuah periode tidak memiliki nilai apa pun (kosong semua), jangan dimasukkan ke JSON output.

⸻

B. RULES: OUTPUT JSON PER ANAK

Output JSON untuk setiap anak harus mengikuti struktur berikut:

{
  "no": <integer>,
  "nik": "<string>",
  "nama_anak": "<string>",
  "tanggal_lahir": "YYYY-MM-DD",
  "jenis_kelamin": "L/P",

  "measurements": [
      {
        "periode": "<string, contoh: 'JANUARI 2024'>",
        "tgl_ukur": "YYYY-MM-DD atau null",
        "umur_bulan": <integer/null>,
        "berat_kg": <float/null>,
        "tinggi_cm": <float/null>,
        "cara_ukur": "<string/null>",

        "status_bb": "<NORMAL / KURANG / LEBIH>",
        "status_tb": "<NORMAL / PENDEK / TINGGI>",

        "rentang_bb_ideal": "<min - max>",
        "rentang_tb_ideal": "<min - max>",

        "status_tb_rasional": "<NORMAL / DANGER / NO_BASELINE / AMBIGU_METHODOLOGY>",
        "catatan_tb_rasional": "<string penjelasan>"
      }
  ]
}

Keterangan:
	•	status_bb dan status_tb menggunakan tabel WHO 0–24 bulan.
	•	status_tb_rasional menilai apakah tinggi badan logis dari bulan sebelumnya.

⸻

C. RULES: KOMPARASI DENGAN TABEL WHO (USIA 0–24 BULAN)
	1.	Cari data WHO berdasarkan umur_bulan.
	2.	Pilih kolom WHO sesuai jenis kelamin:
	•	Jika L: gunakan BB Ideal (L) dan PB Ideal (L)
	•	Jika P: gunakan BB Ideal (P) dan PB Ideal (P)
	3.	Range WHO ditulis dalam format "min-max" (contoh: "5.3-8.8").
Split menjadi:
	•	min_bb
	•	max_bb
	•	min_tb
	•	max_tb
	4.	Penilaian berat (BB):
	•	Jika berat_kg < min_bb → status_bb = "KURANG"
	•	Jika berat_kg > max_bb → status_bb = "LEBIH"
	•	Jika berada dalam range → status_bb = "NORMAL"
	5.	Penilaian tinggi (TB):
	•	Jika tinggi_cm < min_tb → status_tb = "PENDEK"
	•	Jika tinggi_cm > max_tb → status_tb = "TINGGI"
	•	Jika berada dalam range → status_tb = "NORMAL"

⸻

D. RULES: VALIDASI RASIONALITAS TINGGI BADAN PER BULAN

Validasi dilakukan SETIAP PERIODE, dibandingkan dengan periode sebelumnya (jika ada).

Gunakan variabel:
	•	tb_sekarang
	•	tb_sebelumnya
	•	cara_sekarang
	•	cara_sebelumnya

⸻

1. Jika tinggi_bulan_ini < tinggi_bulan_sebelumnya

Jika metode sama:

STATUS = "DANGER"
CATATAN = "Tinggi badan menurun, tidak rasional."

Jika metode berbeda dan selisih < 1 cm:

STATUS = "AMBIGU_METHODOLOGY"
CATATAN = "Penurunan kecil bisa karena beda metode ukur."

Jika metode berbeda dan selisih ≥ 1 cm:

STATUS = "DANGER"
CATATAN = "Penurunan besar, kemungkinan data salah."


⸻

2. Jika tinggi_bulan_sebelumnya KOSONG

STATUS = "NO_BASELINE"
CATATAN = "Tidak ada data sebelumnya untuk memverifikasi rasionalitas tinggi badan."


⸻

3. Jika tinggi_bulan_ini == tinggi_bulan_sebelumnya

STATUS = "NORMAL"
CATATAN = "Tinggi badan stabil dibanding bulan sebelumnya."


⸻

4. Jika tinggi_bulan_ini > tinggi_bulan_sebelumnya

STATUS = "NORMAL"
CATATAN = "Pertumbuhan tinggi badan normal."


⸻

E. CATATAN TAMBAHAN (OPSIONAL TAPI DISARANKAN)
	1.	Perbedaan metode ukur:
	•	TELENTANG biasanya 0.5–1 cm lebih panjang dari BERDIRI.
	•	Maka perbedaan kecil masih wajar.
	2.	Jika usia > 24 bulan dan file WHO tidak mencakup:
	•	berikan status:

status_bb = "OUT_OF_RANGE"
status_tb = "OUT_OF_RANGE"

dan tidak melakukan komparasi WHO.

⸻

💯 Ringkasan Output Akhir

Setiap entry JSON pada measurements[] harus menyertakan:
	•	status gizi berat badan (WHO)
	•	status gizi tinggi badan (WHO)
	•	status rasionalitas tinggi badan antar-periode
	•	catatan penjelasan

