import numpy as np
import pandas as pd
from datetime import time, timedelta


def cal_attendance(fileIn, filenameOut="src/temp/out.xlsx"):
    # %%
    # filenameIn = "สแกนนิ้ว 2024-02.xls"
    # filenameOut = "out_2024_02.xlsx"

    # filenameIn = "แสกนนิ้ว 2024-03.xls"
    # filenameOut = "out_2024_03.xlsx"

    # filenameIn = "แสกนนิ้ว 2024-04.xls"
    # filenameOut = "out_2024_04.xlsx"

    # %%
    dfr = pd.read_excel(fileIn.file)
    fileIn.file.close()

    # %%
    dfr.head()

    # %%
    dfr = dfr.dropna(axis=0, how="all")

    # %%
    locDateCol = dfr.columns.get_loc("Date")
    timeColsOri = dfr.columns.values[locDateCol + 1 :]
    print(timeColsOri)

    # %%
    timeCols = [f"c{i}" for i in range(1, len(timeColsOri) + 1)]
    changeColNameDict = dict(zip(timeColsOri, timeCols))
    dfr = dfr.rename(columns={"ชื่อ-นามสกุล": "name", "Date": "date", **changeColNameDict})
    dfr = dfr[["name", "date", *timeCols]]

    # %%
    dfr.head()

    # %%
    print(dfr.duplicated().sum())
    dfr = dfr[~dfr.duplicated()]
    print(dfr.duplicated().sum())

    # %%
    print(dfr.shape)
    filtNull = dfr[timeCols].isnull().all(axis=1)
    dfr = dfr[~filtNull]
    dfr.shape

    # %%
    def parseDate(dateStr):
        sp = dateStr.split("/")
        day = sp[0]
        month = sp[1]
        year = int(sp[2]) - 543
        return pd.to_datetime(f"{year}/{month}/{day}", format="%Y/%m/%d")

    # Convert to datetime
    dfr["date"] = dfr["date"].apply(parseDate)

    # Convert to date
    dfr["date"] = dfr["date"].dt.date

    # Check
    print(type(dfr["date"].iloc[0]))

    # %%
    # Convert to time
    for col in timeCols:
        dfr[col] = pd.to_datetime(dfr[col], format="%H:%M").dt.time

    # %%
    dfr.head()

    # %%
    # Remove rows with hours outside acceptable range
    def checkTimeOutsideRange(sr):
        return (sr < time(hour=6)) | (sr > time(hour=22))

    filtOutsideRange = pd.Series(data=False, index=dfr.index)
    for col in timeCols:
        _filt = checkTimeOutsideRange(dfr[col])
        filtOutsideRange = filtOutsideRange | _filt
    dfr[filtOutsideRange]

    # %%
    dfr = dfr[~filtOutsideRange]
    print(dfr.shape)

    # %%
    dfr["incompleteInOut"] = False
    filtOneCheckIn = (~dfr[timeCols].isnull()).sum(axis=1) == 1
    dfr.loc[filtOneCheckIn, "incompleteInOut"] = True

    # %%
    # Add time for incomplete check-in/out
    def addTimeForIncompleteCheckInOut(sr):
        dt_time = sr["c1"]
        # Determine whether the missing is the morning in or evening out.
        if dt_time < time(hour=12):  # 12pm
            sr["c2"] = time(hour=17)  # 5pm
        else:
            sr["c2"] = time(hour=8)  # 8am
        return sr

    filtIIO = dfr["incompleteInOut"]
    dfr.loc[filtIIO, :] = dfr.loc[filtIIO, :].apply(
        addTimeForIncompleteCheckInOut, axis=1
    )

    # %%
    dfr.loc[filtIIO].head()

    # %%
    def calculateInOut(row):
        times = row.loc[timeCols].dropna()
        res = times.agg(["min", "max"])
        return pd.concat([row, res])

    dfr = dfr.apply(calculateInOut, axis=1)
    dfr = dfr.rename(columns={"min": "in", "max": "out"})
    dfr.head()

    # %%
    dfr["isInLate"] = dfr["in"] > time(hour=8)
    dfr["isOutEarly"] = dfr["out"] < time(hour=17)
    dfr.head()

    # %%
    # You cannot substract time and time. Need to convert to timedelta first.
    def calInLateMin(dt_time):
        deltaIn = timedelta(hours=dt_time.hour, minutes=dt_time.minute)
        deltaStart = timedelta(hours=8)
        lateMin = (deltaIn - deltaStart).total_seconds() / 60
        if lateMin < 0:
            lateMin = 0
        return lateMin

    dfr["inLateMin"] = dfr["in"].apply(calInLateMin)
    dfr.head()

    # %%
    def calOutEarlyMin(dt_time):
        deltaOut = timedelta(hours=dt_time.hour, minutes=dt_time.minute)
        deltaEnd = timedelta(hours=17)
        earlyMon = (deltaEnd - deltaOut).total_seconds() / 60
        if earlyMon < 0:
            earlyMon = 0
        return earlyMon

    dfr["outEarlyMin"] = dfr["out"].apply(calOutEarlyMin)
    dfr.head()

    # %%
    def calWorkingDuration(row):
        timeIn = row["in"]
        timeOut = row["out"]
        deltaIn = timedelta(hours=timeIn.hour, minutes=timeIn.minute)
        deltaOut = timedelta(hours=timeOut.hour, minutes=timeOut.minute)
        return (deltaOut - deltaIn).total_seconds() / 60

    dfr["workingDuration"] = dfr.apply(calWorkingDuration, axis=1)
    dfr.head()

    # %%
    dfr["overWorkMin"] = dfr["workingDuration"] - (9 * 60)
    dfr.head()

    # %%
    # Calculate working day
    workingDayStart = dfr["date"].min()
    workingDayEnd = dfr["date"].max()
    print(workingDayStart, workingDayEnd)

    # %%
    holidaysThDict = [
        {"date": "2024-01-01", "holiday": "วันขึ้นปีใหม่"},
        {"date": "2024-02-26", "holiday": "วันหยุดชดเชยวันมาฆบูชา"},
        {"date": "2024-04-08", "holiday": "วันหยุดชดเชยวันจักรี"},
        {"date": "2024-04-12", "holiday": "วันหยุดพิเศษ (วันหยุดราชการ)"},
        {"date": "2024-04-15", "holiday": "วันสงกรานต์"},
        {"date": "2024-04-16", "holiday": "วันหยุดชดเชยวันสงกรานต์"},
        {"date": "2024-05-01", "holiday": "วันแรงงานแห่งชาติ (วันหยุดธนาคาร)"},
        {"date": "2024-05-06", "holiday": "วันหยุดชดเชยวันฉัตรมงคล"},
        {"date": "2024-05-22", "holiday": "วันวิสาขบูชา"},
        {"date": "2024-06-03", "holiday": "วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าสุทิดา"},
        {"date": "2024-07-22", "holiday": "วันหยุดชดเชยวันอาสาฬหบูชา"},
        {
            "date": "2024-07-29",
            "holiday": "วันหยุดชดเชยวันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว",
        },
        {"date": "2024-08-12", "holiday": "วันแม่แห่งชาติ"},
        {"date": "2024-10-14", "holiday": "วันหยุดชดเชยวันคล้ายวันสวรรคตรัชกาลที่ 9"},
        {"date": "2024-10-23", "holiday": "วันปิยมหาราช"},
        {
            "date": "2024-12-05",
            "holiday": "วันคล้ายวันพระราชสมภพรัชกาลที่ 9 วันชาติ และ วันพ่อแห่งชาติ",
        },
        {"date": "2024-12-10", "holiday": "วันรัฐธรรมนูญ"},
        {"date": "2024-12-30", "holiday": "วันหยุดพิเศษ (วันหยุดราชการ)"},
        {"date": "2024-12-31", "holiday": "วันสิ้นปี"},
        {"date": "2025-01-01", "holiday": "วันขึ้นปีใหม่"},
        {"date": "2025-02-12", "holiday": "วันมาฆบูชา"},
        {"date": "2025-04-06", "holiday": "วันจักรี"},
        {"date": "2025-04-13", "holiday": "วันสงกรานต์"},
        {"date": "2025-04-14", "holiday": "วันสงกรานต์"},
        {"date": "2025-04-15", "holiday": "วันสงกรานต์"},
        {"date": "2025-04-16", "holiday": "ชดเชยวันสงกรานต์ "},
        {"date": "2025-05-01", "holiday": "วันแรงงานแห่งชาติ "},
        {"date": "2025-05-04", "holiday": "วันฉัตรมงคล"},
        {"date": "2025-05-05", "holiday": "ชดเชยวันฉัตรมงคล"},
        {"date": "2025-05-09", "holiday": "วันพระราชพิธีพืชมงคลจรดพระนังคัลแรกนาขวัญ "},
        {"date": "2025-05-11", "holiday": "วันวิสาขบูชา"},
        {"date": "2025-05-12", "holiday": "ชดเชยวันวิสาขบูชา"},
        {"date": "2025-06-02", "holiday": "พิเศษ"},
        {
            "date": "2025-06-03",
            "holiday": "วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี",
        },
        {"date": "2025-07-10", "holiday": "วันอาสาฬหบูชา"},
        {"date": "2025-07-11", "holiday": "วันเข้าพรรษา "},
        {
            "date": "2025-07-28",
            "holiday": "วันพระบรมราชสมภพ พระบาทสมเด็จพระปรเมนทรรามาธิบดีศรีสินทรมหาวชิราลงกรณ พระวชิรเกล้าเจ้าอยู่หัว",
        },
        {"date": "2025-08-11", "holiday": "พิเศษ"},
        {
            "date": "2025-08-12",
            "holiday": "วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าสิริกิติ์ พระบรมราชินีนาถ พระบรมราชชนนีพันปีหลวง",
        },
        {
            "date": "2025-10-13",
            "holiday": "วันคล้ายวันสวรรคต พระบาทสมเด็จพระบรมชนกาธิเบศร มหาภูมิพลอดุลยเดชมหาราช บรมนาถบพิตร",
        },
        {"date": "2025-10-23", "holiday": "วันปิยมหาราช"},
        {
            "date": "2025-12-05",
            "holiday": "วันคล้ายวันพระบรมราชสมภพ พระบาทสมเด็จพระบรมชนกาธิเบศร มหาภูมิพลอดุลยเดชมหาราช บรมนาถบพิตร",
        },
        {"date": "2025-12-10", "holiday": "วันรัฐธรรมนูญ"},
        {"date": "2025-12-31", "holiday": "วันสิ้นปี"},
    ]

    dfHoliday = pd.DataFrame.from_dict(holidaysThDict)
    dfHoliday["date"] = pd.to_datetime(dfHoliday["date"])
    dfHoliday.head()

    # %%
    # Get working days range
    dateRanges = pd.date_range(start=workingDayStart, end=workingDayEnd)
    dfDateRange = pd.DataFrame(data={"date": dateRanges})
    dfDateRange.head()

    # %%
    def determine_holiday(dt):
        if dt in dfHoliday["date"].values:
            return True
        else:
            return False

    dfDateRange["isHoliday"] = dfDateRange["date"].apply(determine_holiday)
    dfDateRange[dfDateRange["isHoliday"]]

    # %%
    def determine_weekend(dt):
        if dt.weekday() <= 5:  # Monday to Saturday
            return False
        else:
            return True

    dfDateRange["isWeekend"] = dfDateRange["date"].apply(determine_weekend)
    filt = dfDateRange["isHoliday"] | dfDateRange["isWeekend"]
    dfDateRange[filt]

    # %%
    dfDateRange["isWorkingDay"] = ~dfDateRange["isHoliday"] & ~dfDateRange["isWeekend"]

    # %%
    # Convert datetime to date so that I can merge.
    dfDateRange["date"] = dfDateRange["date"].dt.date

    # %%
    def matchDateRange(_dft):
        dft = _dft.copy()
        dft["checkPresent"] = 1
        dfm = pd.merge(
            dfDateRange,
            dft,
            left_on="date",
            right_on="date",
            how="left",
            suffixes=("", "_y"),
        )
        dfm["isPresent"] = dfm["checkPresent"].notnull()
        dfm["isAbsent"] = dfm["checkPresent"].isnull()
        return dfm

    dfg = dfr.groupby(by="name")
    dfgm = dfg.apply(matchDateRange, include_groups=False)

    # Testing
    # dfg = dfr.groupby(by="name")
    # dft = dfg.get_group("รุ้ง").copy()
    # dft["checkPresent"] = 1
    # dfm = pd.merge(dfDateRange, dft, left_on="date", right_on="date", how="left")
    # dfm["isPresent"] = dfm["checkPresent"].notnull()
    # dfm["isAbsent"] = dfm["checkPresent"].isnull()
    # dfm

    # %%
    dfgm = dfgm.reset_index().drop(columns="level_1")

    # %%
    dfgm.columns

    # %%
    dfgm.head()

    # %%
    def determine_is_present_on_working_day(sr):
        isWorkingDay = sr["isWorkingDay"]
        if not isWorkingDay:
            return np.nan
        isPresent = sr["isPresent"]
        if isPresent:
            return True
        else:
            return False

    dfgm["isPresentOnWorkingDay"] = dfgm.apply(
        determine_is_present_on_working_day, axis=1
    )
    dfgm.head()

    # %%
    def determine_is_absent_on_working_day(sr):
        isWorkingDay = sr["isWorkingDay"]
        if not isWorkingDay:
            return np.nan
        isPresent = sr["isPresent"]
        if isPresent:
            return False
        else:
            return True

    dfgm["isAbsentOnWorkingDay"] = dfgm.apply(
        determine_is_absent_on_working_day, axis=1
    )
    dfgm.head()

    # %%
    def determine_is_present_on_holiday_and_weekend(sr):
        isWorkingDay = sr["isWorkingDay"]
        if isWorkingDay:
            return np.nan
        isPresent = sr["isPresent"]
        if isPresent:
            return True
        else:
            return False

    dfgm["isPresentOnHolidayWeekend"] = dfgm.apply(
        determine_is_present_on_holiday_and_weekend, axis=1
    )
    dfgm.head(30)

    # %%
    out1 = (
        dfgm.groupby(by=["name"])
        .agg(
            {
                "isPresent": "sum",
                "isAbsent": "sum",
                "isPresentOnWorkingDay": "sum",
                "isAbsentOnWorkingDay": "sum",
                "isPresentOnHolidayWeekend": "sum",
                "workingDuration": lambda s: s.mean() / 60,
                "overWorkMin": "sum",
                "incompleteInOut": "sum",
                "inLateMin": "sum",
                "outEarlyMin": "sum",
            }
        )
        .rename(
            columns={
                "workingDuration": "workingDuration (mean)",
            }
        )
    )

    # %%
    out2 = dfgm.pivot(index="name", columns="date", values="overWorkMin")

    # %%
    # Using pivot_table, not pivot
    out3 = dfgm.pivot_table(
        index="name", columns="date", values="isPresent", aggfunc="sum"
    )

    # %%
    # Using pivot_table, not pivot
    out4 = dfgm.pivot_table(
        index="name",
        columns="date",
        values="workingDuration",
        aggfunc=lambda s: np.round(s.mean() / 60, 2),
    )

    # %%
    # Write dataframes to Excel with multiple sheets
    names = ["summary", "overWorkMin", "present", "workingDuration (hour)"]
    dataframes = [out1, out2, out3, out4]
    with pd.ExcelWriter(filenameOut) as writer:
        for name, frame in zip(names, dataframes):
            frame.to_excel(writer, sheet_name=name, index=True)
