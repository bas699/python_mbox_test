import mailbox
import sys,io,os,tkinter,tkinter.filedialog,pathlib,shutil,csv,re,logging
from email.header import decode_header
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
logging.basicConfig(level='DEBUG')
logging.info("mboxファイルをテキスト変換します")
logging.info("mboxファイルを選んでください")
root = tkinter.Tk()
root.withdraw()
fTyp =[("mbox","*.mbox"),("","*")]
iDir = os.path.abspath(os.path.dirname(__file__))
file = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
if len(file) ==0:
	logging.error("ファイルを選んでください")
	exit()
mail_box = mailbox.mbox(file)


epath = os.path.splitext(os.path.basename(file))[0]+"export"
w_file = open(epath+".txt",'w',encoding='utf-8')
logging.debug(iDir)
logging.debug(os.path.splitext(os.path.basename(file))[0])
logging.debug(epath)
os.makedirs(epath,exist_ok=True)
os.makedirs(epath+"_temp",exist_ok=True)

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
	w_file.write(str(count)+"\n")
	w_file.write("subject: "+usbj+"\n")
	rfilename=re.sub(r'[\\|\/|:|\*|?|\<|\>|\||"|$|,|;||(\r)|(\n)]+', '', usbj)
	logging.debug(rfilename)
	from_str = a_msg.get_from()
	logging.debug(from_str.split('@')[0])
	w_file0 = open(epath+"/"+from_str.split('@')[0]+"_"+ rfilename + ".txt",'w',encoding='utf-8')
	w_file0.write("subject: "+usbj+"\n")

	w_file.write(from_str+"\n")
	w_file0.write(from_str+"\n")
	try:
		to_str ="From:"+ a_msg.get('from')+"\n"
		to_str += "To:"+a_msg.get('to')+"\n"
		to_str += "Date:"+a_msg.get('date')+"\n"
	except TypeError:
		pass
	logging.debug(to_str)
	w_file.write(to_str+"\n")
	w_file0.write(to_str+"\n")

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
					#logging.error ("** Cannot decode.Cannot specify charset ***"+aa_msg.get("From"))
					a_text = aa_msg.get_payload(decode=True).decode(enc, "ignore")
			w_file.write(a_text[0:400]+"\n")
			w_file0.write(a_text+"\n")
		else:
			attach_e_fname=re.sub(r'[\\|\/|:|\*|?|\<|\>|\||"|$|,|;||(\r)|(\n)|(\t)]+', '', attach_fname)
			with open('./'+epath+"_temp"+"/"+str(count)+"_"+ attach_e_fname, 'wb') as f:
				try:
					f.write(aa_msg.get_payload(decode=True))
				except TypeError:
					pass
				logging.info(attach_fname+"を保存しました。")
		continue
	count += 1
	
	w_file0.close()
w_file.close()
