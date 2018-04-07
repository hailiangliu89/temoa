import sqlite3
import sys, os
import re
import getopt
import xlwt
from xlwt import easyxf 
from collections import defaultdict
from IPython import embed as IP


def make_excel(ifile, ofile, scenario):
	ofile = "db_io/OUTPUT_XLS/"
	tech = defaultdict(list)
	tech_set = set()
	sector = set()
	period = []
	emiss = set()
	row = 0
	count = 0
	sheet = []
	book = []
	book_no = 0
	flag = None
	flag1 = None
	flag2 = None
	i = 0 # Sheet ID
	header = ['Technologies', ]
	header_emiss = []
	header_v = ['Technologies', 'Output Name', 'Vintage', 'Cost']
	tables = {"Output_VFlow_Out" : ["Activity", "vflow_out"], "Output_CapacityByPeriodAndTech" : ["Capacity", "capacity"], "Output_Emissions" : ["Emissions", "emissions"], "Output_Costs" : ["Costs", "output_cost"]}

	
	if ifile is None :
		raise "You did not specify the input file, remember to use '-i' option"
		print "Use as :\n	python DB_to_Excel.py -i <input_file> (Optional -o <output_excel_file_name_only>)\n	Use -h for help."                          
		sys.exit(2)
	else :
		file_type = re.search(r"(\w+)\.(\w+)\b", ifile) # Extract the input filename and extension
	if not file_type :
		print "The file type %s is not recognized. Use a db file." % ifile
		sys.exit(2)
	if ofile is None :
		ofile = file_type.group(1)
		print "Look for output in %s_*.xls" % ofile
		
		
	con = sqlite3.connect(ifile)
	cur = con.cursor()   # a database cursor is a control structure that enables traversal over the records in a database
	con.text_factory = str #this ensures data is explored with the correct UTF-8 encoding

	for k in tables.keys() :
		if not scenario :
			cur.execute("SELECT DISTINCT scenario FROM "+k)
			for val in cur :
				scenario.add(val[0])
		
		for axy in cur.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='technologies';") :
			if axy[0] :
				fields = [ads[1] for ads in cur.execute('PRAGMA table_info(technologies)')]
				if 'sector' in fields :
					cur.execute("SELECT sector FROM technologies")
					for val in cur :
						sector.add(val[0])
					if not sector :
						sector.add('0')
					else :
						flag = 1
		
		if flag is None :
			cur.execute("SELECT DISTINCT tech FROM "+k)
			for val in cur :
				tech['0'].append(val[0])
				tech_set.add(val[0])
		else :
			for x in sector :
				cur.execute("SELECT DISTINCT tech  FROM technologies WHERE sector is '"+x+"'")
				for val in cur :
					if val[0] not in tech[x] :
						tech[x].append(val[0])
						tech_set.add(val[0])
#
		if k is "Output_Emissions" :
			cur.execute("SELECT DISTINCT emissions_comm FROM "+k+" WHERE emissions_comm LIKE 'co2%'")
			for val in cur :
				emiss.add(val[0])
		
		if k is "Output_Costs" :
			pass
		else:#if k is not "Output_V_Capacity":
			cur.execute("SELECT DISTINCT t_periods FROM "+k)
			for val in cur :
				val = str(val[0])
				if val not in period :
					period.append(val)
					header.append(val)
	header[1:].sort()
	period.sort()
	header_emiss = header[:]
	header_emiss.insert(1, "Emission Commodity")
	ostyle = easyxf('alignment: vertical centre, horizontal centre;')
	ostyle_header = easyxf('alignment: vertical centre, horizontal centre, wrap True;')

	for scene in scenario :	
		book.append(xlwt.Workbook(encoding="utf-8"))
		sector={'PowerPlants','supply'}
		for z in sector :
			for a in tables.keys() :
				if z is '0' :
					sheet_name = str(tables[a][0])
					if a is "Output_Costs" :
						flag2 = '1'
					if a is "Output_Emissions" :
						flag1 = '1'
				elif (a is "Output_Costs" and flag2 is None) :
					sheet_name = str(tables[a][0])
					flag2 = '1'
				elif (a is "Output_Emissions" and flag1 is None) :
					sheet_name = str(tables[a][0])
					flag1 = '1'
				elif (a is "Output_Costs" and flag2 is not None) or (a is "Output_Emissions" and flag1 is not None) :
					continue
				else :
					sheet_name = str(tables[a][0])+"_"+str(z)
				sheet.append(book[book_no].add_sheet(sheet_name))
				if a is "Output_Emissions" and flag1 is '1':
					for col in range(0, len(header_emiss)) :
						sheet[i].write(row, col, header_emiss[col], ostyle_header)
						sheet[i].col(col).width_in_pixels = 3300
					row += 1
					for x in tech_set :
						for q in emiss :
							sheet[i].write(row, 0, x, ostyle)
							sheet[i].write(row, 1, q, ostyle)
							for y in period :
								cur.execute("SELECT sum("+tables[a][1]+") FROM "+a+" WHERE t_periods is '"+y+"' and scenario is '"+scene+"' and tech is '"+x+"' and emissions_comm is '"+q+"'")
								xyz = cur.fetchone()
								if xyz[0] is not None :
									sheet[i].write(row, count+2, float(xyz[0]), ostyle)
								else :
									sheet[i].write(row, count+2, '-', ostyle)
								count += 1
							row += 1
							count = 0
					row = 0
					i += 1
					flag1 = '2'
				elif a is "Output_Costs" and flag2 is '1':
					for col in range(0, len(header_v)) :
						sheet[i].write(row, col, header_v[col], ostyle_header)
						sheet[i].col(col).width_in_pixels = 3300
					row += 1
					for x in tech_set :			
						cur.execute("SELECT output_name, vintage, "+tables[a][1]+" FROM "+a+" WHERE scenario is '"+scene+"' and tech is '"+x+"'")
						for xyz in cur :
							if xyz[0] is not None :
								sheet[i].write(row, 0, x, ostyle)
								sheet[i].write(row, count+1, xyz[0], ostyle)
								sheet[i].write(row, count+2, xyz[1], ostyle)
								sheet[i].write(row, count+3, xyz[2], ostyle)
							else :
								sheet[i].write(row, 0, x, ostyle)
								sheet[i].write(row, count+1, '-', ostyle)
								sheet[i].write(row, count+2, '-', ostyle)
								sheet[i].write(row, count+3, '-', ostyle)
							row += 1
						count = 0
					row = 0
					i += 1
					flag2 = '2'
				elif (a is "Output_Costs" and flag2 is '2') or (a is "Output_Emissions" and flag1 is '2'):
					pass
				elif a is not "Output_V_Capacity":
					for col in range(0, len(header)) :
						sheet[i].write(row, col, header[col], ostyle_header)
						sheet[i].col(col).width_in_pixels = 3300
					row += 1
					for x in tech[z] :
						sheet[i].write(row, 0, x, ostyle)
						for y in period :
							cur.execute("SELECT sum("+tables[a][1]+") FROM "+a+" WHERE t_periods is '"+y+"' and scenario is '"+scene+"' and tech is '"+x+"'")
							xyz = cur.fetchone()
							if xyz[0] is not None :
								sheet[i].write(row, count+1, float(xyz[0]), ostyle)
							else :
								sheet[i].write(row, count+1, '-', ostyle)
							count += 1
						row += 1
						count = 0
					row = 0
					i += 1

		if len(scenario) is 1:
			book[book_no].save(ofile+file_type.group(1)+".xls")
		else :
			book[book_no].save(ofile+"_"+scene+".xls")
		
		cur.execute("SELECT t_periods, sector, SUM(emissions) FROM Output_Emissions WHERE emissions_comm LIKE 'co2%' GROUP BY t_periods, sector")
		CO2= cur.fetchall()
		header=['periods','sector','emissions']
		CO2.insert(0,header)
		sheet_name="CO2"
		sheet.append(book[book_no].add_sheet(sheet_name))
		for t, row in enumerate(CO2):
			for j, col in enumerate(row):
				book[book_no].get_sheet(i).write(t,j,col)
		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1

		
		#Commercial total energy consumption
		cur.execute("SELECT t_periods,sector,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE sector='commercial' AND input_comm LIKE 'ELC%' OR sector='commercial' AND input_comm LIKE 'C_NGA_EA%' OR sector='commercial' AND input_comm LIKE 'C_LPG_EA%' OR sector='commercial' AND input_comm LIKE 'C_RFO_EA%' OR sector='commercial' AND input_comm LIKE 'C_DISTOIL_EA%' OR sector='commercial' AND input_comm LIKE 'C_BIO_EA%' GROUP BY t_periods, input_comm")
		Comm= cur.fetchall()
		header=['t_periods','sector','input_comm','SUM(vflow_in)']
		Comm.insert(0,header)
		sheet_name="commercial"
		sheet.append(book[book_no].add_sheet(sheet_name))
		for t, row in enumerate(Comm):
			for j, col in enumerate(row):
				book[book_no].get_sheet(i).write(t,j,col)
		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1

		#Commercial space heating and cooling...
		sheet_name="Comm_SH_SC"
		sheet.append(book[book_no].add_sheet(sheet_name))
		cur.execute("SELECT t_periods,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE tech LIKE 'C_BLND_FUEL_SH%' GROUP BY t_periods, input_comm") #SH fuel mix
		Comm_SH= cur.fetchall()
		header=['t_periods','tech','SUM(vflow_in)']
		Comm_SH.insert(0,header)
		Comm_SH.append(["----","----","----","----","----","----","----","----"]) 
		cur.execute("SELECT t_periods,tech,SUM(vflow_in) FROM Output_VFlow_In WHERE output_comm LIKE 'CSHELC%' GROUP BY t_periods, tech") #ELC for SH by technology
		Comm_SHELC= cur.fetchall()
		header=['t_periods','tech','SUM(vflow_in)']
		Comm_SHELC.insert(0,header)
		Comm_SHELC.append(["----","----","----","----","----","----","----","----"]) 
		cur.execute("SELECT t_periods,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE tech LIKE 'C_BLND_FUEL_SC%' GROUP BY t_periods, input_comm") #ELC for SC by technology
		Comm_SCELC= cur.fetchall()
		header=['t_periods','tech','SUM(vflow_in)']
		Comm_SCELC.insert(0,header)
		Comm_SH_SC=list()
		Comm_SH_SC=Comm_SH+Comm_SHELC+Comm_SCELC
		for t, row in enumerate(Comm_SH_SC):
			for j, col in enumerate(row):
				book[book_no].get_sheet(i).write(t,j,col)

		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1

		#Resinetial total energy consumption
		cur.execute("SELECT t_periods,sector,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE sector='residential' AND input_comm LIKE 'ELC%' OR sector='residential' AND input_comm LIKE 'R_NGA_EA%' OR sector='residential' AND input_comm LIKE 'R_LPG_EA%' OR sector='residential' AND input_comm LIKE 'R_KER_EA%' OR sector='residential' AND input_comm LIKE 'R_DISTOIL_EA%' OR sector='residential' AND input_comm LIKE 'R_BIO_EA%' OR sector='residential' AND input_comm LIKE 'RWHSOL%' GROUP BY t_periods, input_comm")
		Res= cur.fetchall()
		header=['t_periods','sector','input_comm','SUM(vflow_in)']
		Res.insert(0,header)
		sheet_name="residential"
		sheet.append(book[book_no].add_sheet(sheet_name))
		for t, row in enumerate(Res):
			for j, col in enumerate(row):
				book[book_no].get_sheet(i).write(t,j,col)
		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1
		#Residential space heating and cooling...
		sheet_name="Res_SH_SC"
		sheet.append(book[book_no].add_sheet(sheet_name))
		cur.execute("SELECT t_periods,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE tech LIKE 'R_BLND_FUEL_SH%' GROUP BY t_periods, input_comm") #SH fuel mix
		Res_SH= cur.fetchall()
		header=['t_periods','tech','SUM(vflow_in)']
		Res_SH.insert(0,header)
		Res_SH.append(["----","----","----","----","----","----","----","----"]) 
		cur.execute("SELECT t_periods,tech,SUM(vflow_in) FROM Output_VFlow_In WHERE output_comm LIKE 'RSHELC%' GROUP BY t_periods, tech") #ELC for SH by technology
		Res_SHELC= cur.fetchall()
		header=['t_periods','tech','SUM(vflow_in)']
		Res_SHELC.insert(0,header)
		Res_SHELC.append(["----","----","----","----","----","----","----","----"]) 
		cur.execute("SELECT t_periods,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE tech LIKE 'R_BLND_FUEL_SC%' GROUP BY t_periods, input_comm") #ELC for SC by technology
		Res_SCELC= cur.fetchall()
		header=['t_periods','tech','SUM(vflow_in)']
		Res_SCELC.insert(0,header)
		Res_SH_SC=list()
		Res_SH_SC=Res_SH+Res_SHELC+Res_SCELC
		for t, row in enumerate(Res_SH_SC):
			for j, col in enumerate(row):
				book[book_no].get_sheet(i).write(t,j,col)
		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1


		cur.execute("SELECT t_periods,sector,input_comm, SUM(vflow_in) FROM Output_VFlow_In WHERE sector='transport' AND input_comm LIKE 'ELC%' OR sector='transport' AND input_comm LIKE 'DSL_EA%' OR sector='transport' AND input_comm LIKE 'GAS_Z%' OR sector='transport' AND input_comm LIKE 'ETH_CEL_Z%' OR sector='transport' AND input_comm LIKE 'ETH_CORN_Z%' OR sector='transport' AND input_comm LIKE 'H2%' OR sector='transport' AND input_comm LIKE 'CNG_EA%' OR sector='transport' AND input_comm LIKE 'T_LPG_EA%' OR sector='transport' AND input_comm LIKE 'BIODSL_EA%' OR sector='transport' AND input_comm LIKE 'JTF_EA%' OR sector='transport' AND input_comm LIKE 'RFO_EA%' GROUP BY t_periods, input_comm")
		Trn= cur.fetchall()
		header=['t_periods','sector','input_comm','SUM(vflow_in)']
		Trn.insert(0,header)
		sheet_name="transport"
		sheet.append(book[book_no].add_sheet(sheet_name))
		for t, row in enumerate(Trn):
			for j, col in enumerate(row):
				book[book_no].get_sheet(i).write(t,j,col)
		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1

		#LDVs fuel economies
		cur.execute("SELECT t_periods,input_comm, tech,vintage, output_comm, vflow_out FROM Output_VFlow_out WHERE sector='transport' and tech LIKE 'T_LDV_%' and tech!='T_LDV_BLNDDEM'")
		BMT_Out= cur.fetchall()
		cur.execute("SELECT t_periods,input_comm, tech,vintage, output_comm, vflow_in FROM Output_VFlow_in WHERE sector='transport' and tech LIKE 'T_LDV_%' and tech!='T_LDV_BLNDDEM'")
		PJ_In=cur.fetchall()
		sheet_name="FuelEconomy"
		sheet.append(book[book_no].add_sheet(sheet_name))
		header=['t_periods','input_comm','tech','vintage','output_comm','vflow_in','vflow_out']
		for t, row in enumerate(header):
			book[book_no].get_sheet(i).write(0,t,row)
		book[book_no].save(ofile+file_type.group(1)+".xls")
		for t, (row1,row2) in enumerate(zip(PJ_In,BMT_Out)):
			for j, col in enumerate(row1):
				book[book_no].get_sheet(i).write(t+1,j,col)
			book[book_no].get_sheet(i).write(t+1,j+1,row2[-1])
		book[book_no].save(ofile+file_type.group(1)+".xls")
		i += 1
		book_no += 1
		flag1 = None
		flag2 = None
	cur.close()


def get_data(inputs):

	ifile = None
	ofile = "db_io/OUTPUT_XLS/"
	scenario = set()
	
	if inputs is None:
		raise "no arguments found"
		
	for opt, arg in inputs.iteritems():
		if opt in ("-i", "--input"):
			ifile = arg
		elif opt in ("-s", "--scenario"):
			scenario.add(arg)
		elif opt in ("-h", "--help") :
			print "Use as :\n	python DB_to_Excel.py -i <input_file> (Optional -o <output_excel_file_name_only>)\n	Use -h for help."                          
			sys.exit()
	make_excel(ifile, ofile, scenario)

if __name__ == "__main__":	
	
	try:
		argv = sys.argv[1:]
		opts, args = getopt.getopt(argv, "hi:o:s:", ["help", "input=", "output=", "scenario="])
	except getopt.GetoptError:          
		print "Something's Wrong. Use as :\n	python DB_to_Excel.py -i <input_file> (Optional -o <output_excel_file_name_only>)\n	Use -h for help."                          
		sys.exit(2) 
		
	print opts
		
	get_data( dict(opts) )