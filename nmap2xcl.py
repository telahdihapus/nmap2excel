import subprocess
import xlsxwriter
import xlwt
from datetime import datetime
import sys, getopt
import os

#lanjut = raw_input('lanjut ? : ')
now = datetime.now()
nama_file = now.strftime("nmap_%Y%m%d%H%M")+'.xls'
os.system('touch '+str(nama_file))

def buat(ip, mulai, selesai, xa):
	hasil = ['']
	st = 0
	try:
		nama_no = []
		status_port = []
		nama_no.append('ip')
		status_port.append(ip)

		mulai = int(mulai)
		selesai = int(selesai)

		while mulai <= selesai:
			mulai_ = mulai+24
			rang = str(mulai)+'-'+str(mulai_)
			#print(mulai_, selesai)
			sel = selesai - mulai
			if mulai_ > selesai:
				rang = str(mulai)+'-'+str(selesai)
				st = 1
			scan = "nmap -p "+rang+" "+ip
			x = subprocess.getoutput(scan)
			x = str(x)
			x = x.replace('latency).', '%')
			e = 0
			for i in x:
				if i == '%':
					batas = e
				e += 1
			x = x.replace(x[0:batas+2], '')
			x = x.replace('Nmap', '%')
			x = x.replace('seconds', '$')
			e = 0
			for i in x:
				if i == '%':
					batas1 = e
				elif i == '$':
					batas2 = e
				e += 1
			x = x.replace(x[batas1-2:batas2+1], '')
			x = x.replace('PORT', '%')
			x = x.replace('SERVICE', '$')
			e = 0
			for i in x:
				if i == '%':
					batas1 = e
				elif i == '$':
					batas2 = e
				e += 1
			x = x.replace(x[batas1:batas2+2], '')
			s = []
			e = 0
			if '       ' in x:
				x = x.replace('       ', ' ')
			elif '      ' in x:
				x = x.replace('      ', ' ')
			elif '     ' in x:
				x = x.replace('     ', ' ')
			elif '    ' in x:
				x = x.replace('    ', ' ')
			elif '   ' in x:
				x = x.replace('   ', ' ')
			elif '  ' in x:
				x = x.replace('  ', ' ')
			
			if '  ' in x:
				x = x.replace('  ', ' ')
			for i in x:
				if i == '\n':
					s.append(e)
				e += 1
			has = []
			e = 0
			#print(x)
			if st == 0:
				while e < 25:
					if e == 0:
						has.append(x[:s[0]])
					elif e == 24:
						has.append(x[s[e-1]+1:])
					else:
						has.append(x[s[e-1]+1:s[e]])
					e += 1
			else:
				while e <= sel:
					if e == 0:
						has.append(x[:s[0]])
					elif e == sel:
						has.append(x[s[e-1]+1:])
					else:
						has.append(x[s[e-1]+1:s[e]])
					e += 1

			for has_ in has:
				batas = []
				e = 0
				for i in has_:
					if i == ' ':
						batas.append(e)
					e += 1
				no_port = has_[:batas[0]]
				status_port.append(has_[batas[0]+1:batas[1]])
				nama_port = has_[batas[1]+1:]
				nama_no.append(nama_port+'('+no_port+')')
			mulai += 25

		hasil[0] = nama_no
		hasil.append(status_port)
	except KeyboardInterrupt:
		if xa == '':
			exit()
		else:
			judul = 'port'
			wb = xlwt.Workbook()
			ws = wb.add_sheet(judul)

			i = 0
			for xa_ in xa:
				i_ = 0
				for _xa_ in xa_:
					ws.write(i, i_, xa[i][i_])
					i_ += 1
				i += 1

			wb.save(nama_file)
			exit()
	except UnboundLocalError:
		hasil = ':P'
	except TypeError:
		hasil = ':P'
	return hasil 


def pangil(argv):
	ip = 0
	awal = 1
	akhir = 100
	versi = 'versi 1.0'
	try:
		opts, args = getopt.getopt(argv,"i:s:e:ah")
		#print(opts)
		if len(opts) == 0:
			print('silahkah gunakan pilihan nmap2xcl.py -h')
			sys.exit()
	except getopt.GetoptError:
		print('silahkah gunakan pilihan nmap2xcl.py -h')
		sys.exit(2)
	pr = []
	for opt, arg in opts:
		pr.append(opt)
	for opt, arg in opts:
		#print(opt)
		#print(arg)
		if (opt == '-h') or (opt == '--help'):
			print('nmap2xcl.py [pilihan]')
			print('\t-a\ttentang program')
			print('\t-i\tadalah pilihan untuk ip target yang akan discan dengan nmap')
			print('\t-s\tadalah pilihan untuk memasukan awalan range port yang akan discan pada ip')
			print('\t-e\tadalah pilihan untuk memasukan akhiran range port yang akan discan pada ip\n')
			print('contoh penggunaan : nmap2xcl.py -i 192.168.1.1 -s 1 -e 80')
			print('apabila anda tidak memasukan awal dan akhir range port, maka program akan melakukan scan port range 1-100')
			sys.exit()
		elif opt in ("-a"):
			print(open('/root/Documents/python/nmap/banner.txt').read())
			sys.exit()
		elif "-i" not in pr:
			print('tolong masukan ip yang akan discan')
			sys.exit()
		elif opt in ("-i"):
			ip = arg
		elif opt in ("-s"):
			awal = arg
		elif opt in ("-e"):
			akhir = arg
	#print(awal, akhir)
	return versi, ip, awal, akhir

pangil = pangil(sys.argv[1:])
versi = pangil[0]
ip = pangil[1]
awal = pangil[2]
akhir = pangil[3]

print('tunggu sebentar ya, port sedang di scan . . .\n---------------------------')

xa = []
e = 0
xll = buat(ip, awal, akhir, xa)
if xll == ':P':
	print('nmap > '+ip+' > block')
else:
	xl = xll
	print('nmap > '+ip)
	print(str(int(akhir)-int(awal)+1)+' port telah discan')
	print('file anda > '+nama_file)
	if e == 0:
		xa.append(xl[0])
		xa.append(xl[1])
	else:
		xa.append(xl[1])
e += 1

judul = 'port'
wb = xlwt.Workbook()
ws = wb.add_sheet(judul)

i = 0
for xa_ in xa:
	i_ = 0
	for _xa_ in xa_:
		ws.write(i, i_, xa[i][i_])
		i_ += 1
	i += 1

wb.save(nama_file)