sub stockie() 
rowend = cells(rows.count,1).end(xlup).row

dim start_open as Double
dim end_close as Double
dim vol as double
dim counter as integer
dim percent_change as Double
dim yearly_change as double
dim pmin as double
dim pmax as double
dim vmax as double

Cells(1,8).value = "ticker"
cells(1,9).value = "Yearly_Change"
cells(1,10).value= "Percent_Change"
cells(1,11).value = "Total Volume"
cells(1,13).value="Ticker"
cells(1,14).value = "Value"
cells(2,12).value = "Greatest Percent Decrease"
cells(3,12).value = "Greatest Percent Increase"
cells(4,12).value = "Greatest Total Volume"

vol = 0
counter = 2
start_open = cells(2,3).value
for i =2 to rowend
    if cells(i,1).value = cells(i+1,1).value then
        vol= vol+ cells(i,7).value
        if cells(i,3).value = 0 and cells(i+1,3).value <> 0 then
            start_open = cells(i+1,3).value
        end if

    elseif cells(i,1).value <> cells(i+1,1).value then
        cells(counter,8).value = cells(i,1).value
        

        vol = vol+cells(i,7).value
        cells(counter,11).value = vol
        vol = 0

        end_close = cells(i,6).value
        yearly_change = end_close - start_open
        percent_change = (end_close - start_open)/start_open
        cells(counter,9).value = yearly_change
        cells(counter,10).value = percent_change
        
        if cells(i+1,3).value <> 0 then
            start_open = cells(i+1,3)
        end if

        counter= counter +1

    
        
    end if
next i   


pmin = application.worksheetfunction.min(range("j2:j"&counter))
pmax = application.worksheetfunction.max(range("j2:j"&counter))
vmax = application.worksheetfunction.max(range("k2:k"&counter))
for i = 2 to counter
    if cells(i,10).value = pmin then
        cells(2,13).value = cells(i,8).value
        cells(2,14).value = pmin
    elseif cells(i,10).value = pmax then
        cells(3,13).value = cells(i,8).value
        cells(3,14).value = pmax
    elseif cells(i,11).value = vmax then
        cells(4,13).value = cells(i,8).value
        cells(4,14).value = vmax
    end if

    cells(2,14).NumberFormat = "0.00%"
    cells(3,14).NumberFormat = "0.00%"
    cells(i,10).NumberFormat = "0.00%"
    if cells(i,9).value > 0 then
        cells(i,9).interior.colorindex = 4
    else 
        cells(i,9).interior.colorindex = 3   
    end if 
next i


end sub


    