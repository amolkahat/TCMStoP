import xml.etree.ElementTree as ET
import os
import sys
import xlsxwriter

class xmlParse:
	def parseFile(self):
		self.abspath = os.getcwd()+"/"+sys.argv[1]
		tree = ET.parse(self.abspath)
		cols = ['priority','status','author','automated','summary','categoryname','component','defaulttester','notes','testplan_reference','action','expectedresults', 'setup', 'breakdown', 'tag']
		testcases = tree.getroot()
		workbook = xlsxwriter.Workbook('testcases.xlsx')
		worksheet = workbook.add_worksheet()
		rows = 0
		j = 0
		for i in cols:
			worksheet.write(rows,j, i)
			j += 1

		for case in testcases:
			rows = rows + 1
			priority = case.attrib['priority']
			status = case.attrib['status']
			author = case.attrib['author']
			automated = case.attrib['automated']
			summary = ""
			categoryName = ""
			component = ""
			defaulTester = ""
			notes = ""
			testplanReference = ""
			action = ""
			expectedResults = ""
			setup = ""
			breakdown = ""
			tag = ""
			print(priority, status, author, automated)
			for c in case:
				if c.tag == "summary":
					summary = c.text
					if summary is not None:
						summary = summary.replace("<p>","")
						summary = summary.replace("</p>"," ")
					else:
						summary = "None"
				elif c.tag == "categoryname":
					categoryName = c.text
					if categoryName is not None:
						categoryName = categoryName.replace("<p>", "")
						categoryName = categoryName.replace("</p>", "")
					else:
						categoryName = "None"
				elif c.tag == "component":
					component = c.text
					if component is not None:
						component = component.replace("<p>","")
						component = component.replace("</p>","")
						component = component.replace(" ", "")
						component = component.replace("\n", "")
						component.strip()
					else:
						component = "None"
				elif c.tag == "defaulttester":
					defaultTester = c.text
					if defaultTester is not None:
						defaultTester = defaultTester.replace("<p>","")
						defaultTester = defaultTester.replace("</p>","")
					else:
						defaultTester = "None"
				elif c.tag == "notes":
					notes = c.text
					if notes is not None:
						notes = notes.replace("<p>", "")
						notes = notes.replace("</p>", "")
					else:
						notes = "None"
				elif c.tag == "testplan_reference":
					testplanRefernce = c.text
					if testplanReference is not None:
						testplanReference = testplanReference.replace("<p>","")
						testplanReference = testplanReference.replace("</p>","")
					else:
						testplanReference = "None"
				elif c.tag == "action":
					action = c.text
					if action is not None:
						action = action.replace("<p>", "")
						action = action.replace("</p>", "")
					else:
						action = "None"
				elif c.tag == "expectedresults":
					expectedResults = c.text
					if expectedResults is not None:
						expectedResults = expectedResults.replace("<p>", "")
						expectedResults = expectedResults.replace("</p>", "")
					else:
						expectedResults = "None"
				elif c.tag == "setup":
					setup = c.text
					if setup is not None:
						setup = setup.replace("<p>", "")
						setup = setup.replace("</p>", "")
					else:
						setup = "None"
				elif c.tag == "breakdown":
					breakdown = c.text
					if breakdown is not None:
						breakdown = breakdown.replace("<p>", "")
						breakdown = breakdown.replace("</p>", "")
					else:
						breakdown = "None"
				elif c.tag == "tag":
					tag = c.text
					if tag is not None:
						tag = tag.replace("<p>","")
						tag = tag.replace("</p>", "")
					else:
						tag = "None"
				else:
					pass
			print(summary, categoryName, component, defaultTester, notes, testplanReference, action, expectedResults, setup, breakdown, tag)
			worksheet.write(rows,0, priority)
			worksheet.write(rows,1, status)
			worksheet.write(rows,2, author)
			worksheet.write(rows,3, automated)
			worksheet.write(rows,4, summary)
			worksheet.write(rows,5, categoryName)
			worksheet.write(rows,6, component)
			worksheet.write(rows,7, defaultTester)
			worksheet.write(rows,8, notes)
			worksheet.write(rows,9, testplanReference)
			worksheet.write(rows,10, action)
			worksheet.write(rows,11, expectedResults)
			worksheet.write(rows,12, setup)
			worksheet.write(rows,13, breakdown)
			worksheet.write(rows,14, tag)
		workbook.close()
	#		break

if __name__=="__main__":
	if len(sys.argv) == 1:
		print("\n \t Missing file name. \n \t tcmstop.py <file_name> \n")
		sys.exit(1)
	if sys.argv[1].endswith(".xml"):
		if os.path.isfile(os.getcwd()+"/"+sys.argv[1]):
			print("\n \t File found.")
			xmlP = xmlParse()
			xmlP.parseFile()	
		else:
			print("\n \t File not found.")
			sys.exit(1)
	else:
		print("\n \t Specify XML file.")
		sys.exit(1)
