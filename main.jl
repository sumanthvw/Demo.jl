using DataFrames, XLSX

client_needs = DataFrame(XLSX.readtable("hs.contractbranchclientneedslevel.xlsx", "Sheet1", first_row=5, stop_in_empty_row=false))

oldNames = names(client_needs)
newNames = Vector{String}(undef, length(oldNames))
for (i, name) in enumerate(oldNames)
    new = replace(name, " " => "")
    newNames[i] = new
end

rename!(client_needs, newNames)

client_needs = dropmissing(client_needs, :HSOFFICE)

akl1 = filter(row -> row.HSOFFICE == "AKL", client_needs)
wai1 = filter(row -> row.HSOFFICE == "HAM" || row.HSOFFICE == "PAE", client_needs)
bop1 = filter(row -> row.HSOFFICE == "BOP" || row.HSOFFICE == "TGA", client_needs)

client_active = DataFrame(XLSX.readtable("hsclientclientactivelist.xlsx", "Sheet1", first_row=5, stop_in_empty_row=false))

oldNames = names(client_active)
newNames = Vector{String}(undef, length(oldNames))
for (i, name) in enumerate(oldNames)
    new = replace(name, " " => "")
    newNames[i] = new
end

rename!(client_active, newNames)

client_active = dropmissing(client_active, :HSOFFICE)

akl2 = filter(row -> row.HSOFFICE == "AKL", client_active)
wai2 = filter(row -> row.HSOFFICE == "HAM" || row.HSOFFICE == "PAE", client_active)
bop2 = filter(row -> row.HSOFFICE == "BOP" || row.HSOFFICE == "TGA", client_active)

employee_appointments = DataFrame(XLSX.readtable("hsapptemployeeappointmentlist.xlsx", "Sheet1", first_row=5, stop_in_empty_row=false))

oldNames = names(employee_appointments)
newNames = Vector{String}(undef, length(oldNames))
for (i, name) in enumerate(oldNames)
    new = replace(name, " " => "")
    newNames[i] = new
end

rename!(employee_appointments, newNames)

employee_appointments = dropmissing(employee_appointments, :OFFICE)

akl3 = filter(row -> row.OFFICE == "AKL", employee_appointments)
wai3 = filter(row -> row.OFFICE == "HAM" || row.OFFICE == "PAE", employee_appointments)
bop3 = filter(row -> row.OFFICE == "BOP" || row.OFFICE == "TGA", employee_appointments)

employee_address = DataFrame(XLSX.readtable("payemployeeaddress.xlsx", "Sheet1", first_row=7, stop_in_empty_row=false))

oldNames = names(employee_address)
newNames = Vector{String}(undef, length(oldNames))
for (i, name) in enumerate(oldNames)
    new = replace(name, " " => "")
    newNames[i] = new
end

rename!(employee_address, newNames)

employee_address = dropmissing(employee_address, :OFFICE)

akl4 = filter(row -> row.OFFICE == "AKL", employee_address)
wai4 = filter(row -> row.OFFICE == "HAM" || row.OFFICE == "PAE", employee_address)
bop4 = filter(row -> row.OFFICE == "BOP" || row.OFFICE == "TGA", employee_address)

employee_skills = DataFrame(XLSX.readtable("payemployeeskillslist.xlsx", "Sheet1", first_row=9, stop_in_empty_row=false))

oldNames = names(employee_skills)
newNames = Vector{String}(undef, length(oldNames))
for (i, name) in enumerate(oldNames)
    new = replace(name, " " => "")
    newNames[i] = new
end

rename!(employee_skills, newNames)

employee_skills = DataFrame([ffill(employee_skills[!, HOMEOFFICE]) for HOMEOFFICE in names(employee_skills)], names(employee_skills))

akl5 = filter(row -> row.HOMEOFFICE == "AKL", employee_skills)
wai5 = filter(row -> row.HOMEOFFICE == "HAM" || row.HOMEOFFICE == "PAE", employee_skills)
bop5 = filter(row -> row.HOMEOFFICE == "BOP" || row.HOMEOFFICE == "TGA", employee_skills)

client_appointments = DataFrame(XLSX.readtable("hsapptclientappointmentlist.xlsx", "Sheet1", first_row=5, stop_in_empty_row=false))

oldNames = names(client_appointments)
newNames = Vector{String}(undef, length(oldNames))
for (i, name) in enumerate(oldNames)
    new = replace(name, " " => "")
    newNames[i] = new
end

rename!(client_appointments, newNames)

client_appointments = dropmissing(client_appointments, :OFFICE)

akl6 = filter(row -> row.OFFICE == "AKL", client_appointments)
wai6 = filter(row -> row.OFFICE == "HAM" || row.OFFICE == "PAE", client_appointments)
bop6 = filter(row -> row.OFFICE == "BOP" || row.OFFICE == "TGA", client_appointments)

XLSX.writetable("AKL.xlsx", "hs.contractbranchclientneedslev" => akl1, 
                            "hsclientclientactivelist" => akl2,
                            "hsapptemployeeappointmentlist" => akl3,
                            "payemployeeaddress" => akl4,
                            "payemployeeskillslist" => akl5,
                            "hsapptclientappointmentlist" => akl6)

XLSX.writetable("BOP.xlsx", "hs.contractbranchclientneedslev" => bop1, 
                            "hsclientclientactivelist" => bop2,
                            "hsapptemployeeappointmentlist" => bop3,
                            "payemployeeaddress" => bop4,
                            "payemployeeskillslist" => bop5,
                            "hsapptclientappointmentlist" => bop6)

XLSX.writetable("WAI.xlsx", "hs.contractbranchclientneedslev" => wai1, 
                            "hsclientclientactivelist" => wai2,
                            "hsapptemployeeappointmentlist" => wai3,
                            "payemployeeaddress" => wai4,
                            "payemployeeskillslist" => wai5,
                            "hsapptclientappointmentlist" => wai6)

