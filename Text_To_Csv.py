# On Time Delivery Details - Excel
# Brandon Sturgeon

import csv
import re
import sys


header = ["Item ID",
          "Planner Name", 
          "PO Number",
          "PO Line Number",
          "Qty Needed",
          "Qty Received", 
          "Receipt $s",
          "ASN Qty Shipped",
          "Expected Ship Date",
          "ASN/ Receipt Date",
          "Days Late(-) Early(+)",
          "ASN Data"]

# Opens the .csv (Creates if it doesn't exist) and writes the data
def send_to_spreadsheet(argv):
    # Opens file and creates headers
    def sendback():
        print "Click and drag a properly formatted .txt onto the .exe"
        return

    if argv[1][-4:] != ".txt":
        sendback()
        
    else:
        filename = str(argv[1][:-4])
        with open(filename+".csv", "wb") as f:
            the_writer = csv.writer(f)
            the_writer.writerows([header])
            content = []
            # Opens the text file to extract data
            with open(argv[1], "rb") as r:
                read = r.read()
                table_data = re.findall("Data\r\n(.+?)Supplier", read, re.DOTALL)
                for l in table_data:
                    l = re.findall("(.+?)\s(Y|N)", l, re.DOTALL)
                    for b in l:
                        # Modifies string to better send to Excel
                        b = " ".join(b)
                        b = b.replace("\r\n"," ")
                        b = b.strip()

                        find_iter = [m.start() for m in re.finditer(" ", b)][1]
                        b1 = b[:find_iter].strip(), b[find_iter:].strip()
                        b = "_".join(b1)
                        content.append(b)
                        
            the_writer.writerows([l.split(" ") for l in content])

if __name__ == "__main__":
    for arg in sys.argv:
        print arg
    sys.exit(send_to_spreadsheet(sys.argv))








            
