from posixpath import split
import xlsxwriter
import universities
uni = universities.API()


def main():
	# Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook('proje.xlsx')
	worksheet = workbook.add_worksheet()

	lines="""robinson.noble@duke.edu
	jenni.railevirta@utu.fi
	christian.troesch@unisg.ch
	arpitasurendra.halgeri@nus.edu.sg
	adam.prime@utoronto.ca
	Loic.SADOULET@insead.edu
	jeffrey.scott@colorado.edu
	Jessica.Packard@colorado.edu
	sabrena.robinson@duke.edu
	clark.wray@colorado.edu
	luis.marquina@duke.edu
	Joel.Goldstein@bristol.ac.uk
	ole.langfeldt@ntnu.no
	koen.schoors@ugent.be
	andre.simoneau@mcgill.ca
	gerald.crawford@oberlin.edu
	erika.farnstrand@hhs.se
	dolf.jordaan@up.ac.za
	dave.wilson@rmit.edu.au
	hong.peng@yale.edu
	arenno.mba2006@london.edu
	Jeff.Sieracki@colorado.edu
	wolfgang.klas@univie.ac.at
	juha.kotimaa@helsinki.fi
	auke.ijspeert@epfl.ch
	Lisa_Beisser@unc.edu
	andy.nix@bristol.ac.uk
	lucilla.poston@kcl.ac.uk"""

	lines_splitted = [i.strip("\t") for i in lines.splitlines()]

	with open("emailsandnamesUntitled-2.txt","r+") as f:
		lines_splitted = [i.strip("\t") for i in f.readlines()]

	toplam_satir = len(lines_splitted)
	# Start from the first cell. Rows and columns are zero indexed.

	worksheet.write(0,0,"E-posta adresi")
	worksheet.write(0,1,"İsim Soyisim")
	worksheet.write(0,2,"Üniversite")
	worksheet.write(0,3,"Ülke")

	row = 1
	col = 0

	# Iterate over the data and write it out row by row.

	

	for i,email in enumerate(lines_splitted):
		print(str(i)+"/"+str(toplam_satir))
		worksheet.write(row, col,     email)
		ad_soyad,universite,ulke = ayır(email)
		worksheet.write(row, col + 1, f"{ad_soyad[0][0]} {ad_soyad[-1][0]}" )
		worksheet.write(row, col + 2, f"{universite}" )
		worksheet.write(row, col + 3, f"{ulke}" )
		row += 1
		


	workbook.close()
def ayır(email:str):

	splt_email = email.split("@")
	universite0 = splt_email[-1].splitlines()
	
	universite = uni.search(domain = universite0)
	pArray = [nn for nn in universite]
	try:
		universite = pArray[0].name
	except:
		universite = splt_email[-1]+"HATAUNI"
	
	try:
		ulke= pArray[0].country
	except:
		ulke= splt_email[-1]+"HATAULKE"

	splt_0 =	splt_email[0].split(".")
	
	splt_1 = [i.split("_") for i in splt_0]
	
	splt_2 = [i[0].split("-") for i in splt_1]
	
	
	return splt_2,universite,ulke
main()