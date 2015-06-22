files=Dir.entries(".")
files.each do |file|
  if File.extname(file)==".ps" then
    puts file
    filename=File.basename(file,".ps")+".pdf"
    puts filename
    modfile = File.open(file, "r")
    found = false
    addr = " "
    regcode = " "
    subject = " "
    modfile.each do |line|
        if line =~ /Subject \((.*)\)/
         regcode = $1
         puts regcode
        end
    
        if line =~ /Title \((.*)\)/
          subject = $1
          puts subject
        end
    
    
      found = found || line.include?("Your e-mail address:")
      if found then
        if line =~ /\((.*@.*)\)/
         addr = $1.split("; ")
         addr = addr.map {|elem| elem.gsub(",","")}
         puts addr
        end
      end
    end
    subject = regcode + " - " + subject
    puts subject


  end
end