fout = open('findbook.csv','w',encoding='utf-8')
fin = open('allwordtext','r',encoding='utf-8')

for line in fin:
	#if re.search(gkey, line) :    
	if line.find('.docx') != -1:
		fout.write('**************************************\n')
		fout.write(line)
		fname = line.strip()
	if line.find('相关文件') != -1:
		fout.write(fname)
		fout.write(',')	
		fout.write(line)
	start1 = '《'
	end1 = '》'
	s = line.find(start1)
	while s!=-1:
		e = line.find(end1, s)
		sub_str = line[s:e + len(end1)]
		fout.write(fname)
		fout.write(',')	
		fout.write(sub_str)	
		fout.write('\n')	
		s = line.find(start1, e)
		
fout.close()
fin.close()