使用說明:
1. 將考勤機中取出的檔案放到attendance目錄
	如果考勤就跨月份, 請改檔名再放入
	例如考勤跨3月與4月:
		3考勤報表.xls
		4考勤報表.xls

2. 打開settings.ini並填入正確的時間和檔案名稱
	範例1 (沒有跨月份, 即所有考勤在同一月份):
	start_year=2023
	start_mon=4
	start_day=3
	start_attendance_filename=4考勤報表.xls

	end_year=2023
	end_mon=4
	end_day=9
	end_attendance_filename=4考勤報表.xls
	===========================================================
	範例2 (跨月份, 即所考勤分別在兩個月份):
	start_year=2023
	start_mon=3
	start_day=27
	start_attendance_filename=3考勤報表.xls

	end_year=2023
	end_mon=4
	end_day=2
	end_attendance_filename=4考勤報表.xls

3. 執行aj_frog.exe, 執行完後結果會放在result目錄中, 檔名開頭是執行時的日期與時間, 方便確認

軟體工作原理:
1. 根據settings.ini裡的設定讀取attendance目錄中的檔案
2. 根據settings.ini裡的開始與結束日期讀取考勤檔中的時間
3. 將時間填入template\考勤彙總報表.xlsx中, 並另存到result目錄中
