time=Time.new
puts time.year
puts time.month
puts time.day

date=time.strftime('%Y-%m-%d')
puts date

newname="convert_ps2pdf."+date

File.rename("convert_ps2pdf.bat", newname)
modfile = File.open("convert_ps2pdf.bat", "w")

modfile.puts("echo on")

modfile.puts()
modfile.puts('cd D:\studies\childhood\QUESTIONNAIRES\ps_files')
modfile.puts()

files=Dir.entries(".")
files.each do |file|

  newfile=file.sub('.ps','.pdf')

  if newfile!=file then
    begin
      modfile.write '"C:\Program Files\gs\gs9.15\bin\gswin64c" -sDEVICE=pdfwrite -o '
      modfile.write newfile
      modfile.write ' '
      modfile.puts(file)
    end
  end
end

modfile.close

system('convert_ps2pdf.bat') 
