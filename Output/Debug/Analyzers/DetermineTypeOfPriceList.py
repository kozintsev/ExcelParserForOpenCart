import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

def determine_type(range):
    result = 1
    s = range["C2"]
	if s.find("Два Союза") >= 0:
		return "For2Union.py"
    return ""