require 'win32ole'


receipients = ["masuyere@iarc.fr"]
subjectline = "IICC-3 questionnaire"
emailtext   = "Dear Colleague,\n\nPlease find attach a pdf copy of the questionnaire that you completed yesterday.\n\nBest regards,\n\nEric Masuyer\n\n\nEric Masuyer\nData Manager\nSection of Cancer Information\n\nInternational Agency for Research on Cancer (IARC/WHO)\n150 cours Albert-Thomas\n69372 Lyon Cedex 08\nFrance"

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
    message.Attachments.Add('C:\studies\childhood\QUESTIONNAIRES\ps_files\ALARGNEU_20140225.ps',1)
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