import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

def analyze(range):
    i = 1
    result = 1
	s = range["D10"]
	result = s
    # while True:
		  # s1 = range["A" + i]
      # s2 = range["B" + i]
      # if (s1 == '' and s2 == '') break
		  # i = i + 1

    return result