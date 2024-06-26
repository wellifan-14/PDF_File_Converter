# PDF_File_Converter
Aplikasi PDF File Converter berbasis desktop dibangun menggunakan Python. Aplikasi ini dirancang untuk memudahkan pengguna dalam mengonversi berbagai jenis file menjadi format PDF. Dalam pengembangannya, beberapa library Python digunakan untuk memastikan fungsionalitas dan antarmuka pengguna yang intuitif.

- Tkinter: Digunakan untuk membuat antarmuka grafis (GUI) aplikasi. Dengan Tkinter, pengguna dapat dengan mudah berinteraksi dengan aplikasi melalui tombol dan jendela yang user-friendly.

- PIL (Python Imaging Library): Digunakan untuk mengonversi gambar ke format PDF. PIL memungkinkan aplikasi untuk membuka, memanipulasi, dan menyimpan berbagai format gambar.

- win32com: Digunakan untuk mengonversi file Microsoft Word dan PowerPoint ke format PDF. Library ini memanfaatkan fungsi COM (Component Object Model) di Windows untuk mengotomatisasi proses konversi dokumen.

- os: Digunakan untuk operasi sistem seperti mengelola jalur file dan nama file selama proses konversi.

Setelah aplikasi dikembangkan, file Python (.py) yang berisi kode aplikasi dikonversi menjadi file executable (.exe) menggunakan library py_to_exe. Konversi ini memungkinkan aplikasi untuk dijalankan pada sistem Windows tanpa memerlukan instalasi Python, membuat distribusi dan penggunaan aplikasi lebih mudah bagi pengguna akhir.

Fitur utama dari aplikasi ini meliputi:

- Konversi Word ke PDF: Pengguna dapat memilih file .docx dan mengonversinya menjadi PDF dengan satu klik.
- Konversi PowerPoint ke PDF: File presentasi .pptx dapat dikonversi menjadi PDF untuk distribusi yang lebih mudah.
- Konversi Gambar ke PDF: Aplikasi ini mendukung konversi berbagai format gambar (seperti .jpg, .jpeg, .png) ke PDF.
