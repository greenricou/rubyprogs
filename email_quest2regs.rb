require 'win32ole'


receipients = ["masuyere@iarc.fr","masuyer@iarc.fr"]

emailtext   = "Dear Colleague,\n\nPlease find attach a pdf copy of the questionnaire that you completed the 18th June 2015.\n\nBest regards,\n\nEric Masuyer\n\n\nEric Masuyer\nData Manager\nSection of Cancer Surveillance (CSU)\n\nInternational Agency for Research on Cancer (IARC/WHO)\n150 cours Albert Thomas\n69372 Lyon Cedex 08\nFrance"

files=Dir.entries(".")
files.each do |file|
  if File.extname(file)==".ps" then
    puts file
    modfile = File.open(file, "r")
    found = false
    addr = " "
    addrs = []
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
         addrs << addr
        end
      end
    end
    subjectline = regcode + " - " + subject
    puts subjectline

    path=File.expand_path(file)
    puts path
    filename=path.sub('.ps','.pdf')
    puts filename
    addrs << "ccs@iarc.fr"
    receipients=addrs.flatten
    puts receipients



begin
  outlook = WIN32OLE.new('Outlook.Application')
rescue Exception=>e
  puts "No outlook found"
  outlook = nil
end
if outlook != nil
  receipients.each do |email|

    message         = outlook.CreateItem(0)
    message.To      = email
    message.Subject = subjectline
    message.Body    = emailtext
    message.Attachments.Add(filename, 1)
  #  These lines are if you want to attach files...
  #  cal_files.each do |file_name|
  #    message.Attachments.Add(file_name, 1)
  #  end
  #  1.upto(message.attachments.count) do |index|
  #    message.attachments(index).PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001E",  "text/calendar")
  #  end
    message.Send # unless DEBUG
  end
end
end
end
