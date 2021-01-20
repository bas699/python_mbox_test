import mailbox
import sys,io,os,tkinter,tkinter.filedialog,pathlib,shutil,csv,re
from email.header import decode_header
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
root = tkinter.Tk()
root.withdraw()
fTyp =[("mbox","*.mbox"),("","*")]
iDir = os.path.abspath(os.path.dirname(__file__))
file = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
mail_box = mailbox.mbox(file)

w_file = open("export.txt",'w',encoding='utf-8')
count = 0
for key in mail_box.keys():
	a_msg = mail_box.get(key)

	usbj = ''
	if a_msg['Subject'] != None:
		for bstr,enc in decode_header(a_msg['Subject']) :
			if enc == None:
				if isinstance(bstr,bytes):
					usbj += bstr.decode("ascii", "ignore")
				else:
					usbj += bstr
			elif enc == "unknown-8bit":#shift-jis
				usbj += bstr.decode("cp932", "ignore")
			else:
				if isinstance(bstr,bytes):
					usbj += bstr.decode(enc, "ignore")
				else:
					usbj += bstr
			usbj += ","
	w_file.write("subject: "+usbj+"\n")
	rfilename=re.sub(r'[\\|\/|:|\*|?|\<|\>|\||"|$|,|;||(\r)|(\n)]+', '', usbj)
	print(rfilename)
	w_file0 = open(str(count)+"_"+ rfilename + ".txt",'w',encoding='utf-8')
	w_file0.write("subject: "+usbj+"\n")
	from_str = a_msg.get_from()
	w_file.write(from_str+"\n")
	w_file0.write(from_str+"\n")
	#to_str = a_msg.get_to()
	#print(to_str)
	for aa_msg in a_msg.walk():
		if  'multipart' in aa_msg.get_content_type():
			continue #"text"パートでなかったら次のパートへ
			attach_fname = aa_msg.get_filename()
			if not attach_fname:
				if aa_msg.get_content_charset() :
					a_text = aa_msg.get_payload(decode=True).decode(aa_msg.get_content_charset(), "ignore")
				else:
					if "charset=shift_jis" in str(aa_msg.get_payload(decode=True)):
						a_text = aa_msg.get_payload(decode=True).decode("cp932", "ignore")
					elif enc == None:#ascii
						a_text = aa_msg.get_payload(decode=True).decode("ascii", "ignore")
					else:
						#print ("** Cannot decode.Cannot specify charset ***"+aa_msg.get("From"))
						a_text = aa_msg.get_payload(decode=True).decode(enc, "ignore")
				w_file.write(a_text+"\n")
				w_file0.write(a_text+"\n")
			else:
				with open('./' + attach_fname, 'wb') as f:
					f.write(aa_msg.get_payload(decode=True))
					print(f"{filename}を保存しました。")
		continue
	count += 1
	w_file0.close()
w_file.close()
