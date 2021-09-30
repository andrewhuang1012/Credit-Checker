# Credit-Checker
data資料夾:只放電算中心所提供的學生成績excel，格式如原本既有的檔案。 該資料夾僅能放一個excel，否則只會執行其中之一個檔案。

Rules資料夾：
共有3個excel檔案：  
course_table.xlsx：該檔案為電子系以及電資學院目前既有的課程配分表，若有新增則手動加入於表格尾端。  
Capstone.xlsx：該檔案為Capstone課程，即修課紀錄將顯示於成績單最下面的課程(實務專題、電子總整課程等)。  
Specialcourse.xlsx：此為特殊課程，暑期校外實習因時數不同獲得學分數也不同，所以將"課程名稱"輸入於此excel表格即能讓程式自動讀取學生的修課學分數，無須例外處理(建配分表)。  

out資料夾：存放程式執行後的成績單，若不符合規範要求，檔案尾端會出現fail(如B10302000_Fail.xls)

representatives資料夾：該資料夾存放寫入報告書的代表，每班高中低各2位(代號:高H，中M，低:B)。

執行方式：
先點兩下執行excel_writer.exe再執行Representatives2.exe

excel_writer.exe會顯示Welcome NTUST ECE Credits System，執行後成績單會存放於out資料夾，會輸出通過人數、成功率等數據。
Representative.exe會顯示一長串列表，此列表為(代表,排名)，執行後會存放於representatives

注意事項：
1.資料夾名稱請勿更改
2.請檢察course_table.xlsx檔案內的配比後再執行。
