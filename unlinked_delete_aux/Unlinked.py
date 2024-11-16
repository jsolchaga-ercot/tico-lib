

file = open('C://Users//JSOLCHAGA//OneDrive - ERCOT//Documents//Python//pyrtp//pyrtp//RTP_Tools//small_useful_scripts//unlinked_delete_aux//Unlinked.txt', 'r')
out_file = open('C://Users//JSOLCHAGA//OneDrive - ERCOT//Documents//Python//pyrtp//pyrtp//RTP_Tools//small_useful_scripts//unlinked_delete_aux//Backout.aux', 'w')


out_file.write("// Delete broken contingencies.\n")
out_file.write("SCRIPT\n")
out_file.write('{\n')
for line in file:
    line = line.replace('\n', '')
    out_file.write("Delete(Contingency,\"<DEVICE>Contingency \'" + line + "\' \");\n")
out_file.write('}')
out_file.close()
