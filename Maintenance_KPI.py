import sys
sys.path.append("..") # Adds higher directory to python modules path.
import pyodbc
from datetime import date, timedelta, datetime
import time
import datetime
from decimal import Decimal
import xlsxwriter
from emailing import emailing
import os
import reporting

# Set the start date - this report is expected to be run on Tuesdays
startDate = date.today() - timedelta(days = 1)

# Set the timestamp
ts = time.time()
st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')

# email subject
eSubject = 'Maintenance KPI for ' + str(startDate) 

# creation of the excel workbook
writer = xlsxwriter.Workbook('Maintenance_KPI_' + str(startDate) + '.xlsx') 

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server={REDACTED};'
                      'Database={REDACTED};'
                      'Trusted_Connection=yes;')

#Declare Cursor
cursor = conn.cursor() 

# sql to get the date of the sunday of this business week last year
dateSql = """DECLARE @startDate date, @year int, @week int

             select @year = year - 1, @week = WeekOfYear
             from {REDACTED} where date = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

             SET @startDate = (select min(date) 
             from {REDACTED} where year = @year and WeekOfYear = @week)

             select @startDate
             """

cursor.execute(dateSql)

rollingYearBegin = cursor.fetchone()

# contains the beginning date of the range needed for the rolling 52 week line charts
lastYear = str(rollingYearBegin[0]) 

# sql to get each year-week pairing through this business week last year
weekSql = """
    DECLARE @startDate date, @endDate date, @year int, @week int

    select @year = year - 1, @week = WeekOfYear
    from {REDACTED} where date = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

    SET @startDate = (select min(date) 
    from {REDACTED} where year = @year and WeekOfYear = @week)

    SET @endDate = (select date
    from {REDACTED} where date = cast(dateadd(day, -2, '""" + str(startDate) + """') as date))

    select distinct year, WeekOfYear
    from {REDACTED}
    where date between @startDate and @endDate
    order by year desc, WeekOfYear desc
"""

cursor.execute(weekSql)

### contains 53-entry list of year/week combos as of the date of running.
### Explanation: for each row, [0] is year and [1] is week, ordered descending. So row[0] = [2020, 44] and row[52] = [2019, 44]
weeks = cursor.fetchall()

### begin methods ###

#### data collecting methods ####

# to get accident total cost for one company slice for one week ({REDACTED} 12, 31, 32, 36, 37, and 39, {REDACTED} 000023 and 000025, and {REDACTED} 71 associated with tire {REDACTED})
def getAcc(week, comp):
    acc = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
        where ({REDACTED} in ('0031', '0032', '0036', '0037', '0039', '0012')
        or {REDACTED} = '0021'
        or ({REDACTED} = '000023' or {REDACTED} = '000023' or {REDACTED} = '000025' or {REDACTED} = '000025')
        or (({REDACTED} = '017000' or {REDACTED} = '017000') and {REDACTED} = '71'))
        and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        and {REDACTED} not like 'W%'
        and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    acc = cursor.fetchone()

    if acc[0] is None:
        return 0
    else:
        return acc[0]

# to get tires total cost for one company slice for one week ({REDACTED} starting with 017, not including labor or accident costs)
def getTires(week, comp):
    tires = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join Dim.{REDACTED} on {REDACTED} = date
        where {REDACTED} = 'PT'
        and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        and {REDACTED} not like 'W%'
        and ({REDACTED} like '017%' or {REDACTED} like '017%')
        and not ({REDACTED} in ('0031', '0032', '0036', '0037', '0039', '0012')
        or {REDACTED} = '0021'
        or ({REDACTED} = '000023' or {REDACTED} = '000023' or {REDACTED} = '000025' or {REDACTED} = '000025')
        or (({REDACTED} = '017000' or {REDACTED} = '017000') and {REDACTED} = '71'))
        and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    tires = cursor.fetchone()

    if tires[0] is None:
        return 0
    else:
        return tires[0]

# to get outside repair costs for one company slice for one week ({REDACTED} other than 0001, not including accident costs or tire part cost)
def getORO(week, comp):
    oro = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
        where {REDACTED} != '0001'
        and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        and {REDACTED} not like 'W%'
        and not (({REDACTED} like '017%' or {REDACTED} like '017%') and {REDACTED} = 'PT')
        and not ({REDACTED} in ('0031', '0032', '0036', '0037', '0039', '0012')
        or {REDACTED} = '0021'
        or ({REDACTED} = '000023' or {REDACTED} = '000023' or {REDACTED} = '000025' or {REDACTED} = '000025')
        or (({REDACTED} = '017000' or {REDACTED} = '017000') and {REDACTED} = '71'))
        and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    oro = cursor.fetchone()

    if oro[0] is None:
        return 0
    else:
        return oro[0]

# to get all non-indirect costs not covered by previous sql statements for one company slice for one week (everything in {REDACTED} 0001, not including accident, tire, or ORO ROs)
def getMntRep(week, comp):
    mntRep = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
        where {REDACTED} = '0001'
        and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        and {REDACTED} not like 'W%'
        and not (({REDACTED} like '017%' or {REDACTED} like '017%') and {REDACTED} = 'PT')
        and not ({REDACTED} in ('0031', '0032', '0036', '0037', '0039', '0012')
        or {REDACTED} = '0021'
        or ({REDACTED} = '000023' or {REDACTED} = '000023' or {REDACTED} = '000025' or {REDACTED} = '000025')
        or (({REDACTED} = '017000' or {REDACTED} = '017000') and {REDACTED} = '71'))
        and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    mntRep = cursor.fetchone()

    if mntRep[0] is None:
        return 0 
    else:
        return mntRep[0]

# to get all non-RO labor costs for one company slice for one week (all indirect charges that aren't parts, taxes, or fees)
def getIndirectCosts(week, comp):
    indCost = 0

    sql = """
        SELECT sum({REDACTED})
        FROM {REDACTED} ic
        join {REDACTED} dd on dd.date = ic.date
        where concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        and {REDACTED} = 'L'
        and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    indCost = cursor.fetchone()

    if indCost[0] is None:
        return 0 
    else:
        return indCost[0]

# to get total cost of repairs on a single manufacturing year of vehicle for one company slice for one week
def getMFGYear(week, year, comp):
    mfg = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
        join {REDACTED} on {REDACTED} = {REDACTED}
        where {REDACTED} = '""" + str(year) + """'
        and {REDACTED} not like 'W%'
        and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    mfg = cursor.fetchone()

    if mfg[0] is None:
        return 0
    else:
        return mfg[0]

# to get repair costs associated with driver-faulted accidents for one company slice for one week ({REDACTED} 0012 or 10 {REDACTED} for 058010 {REDACTED})
def getDriverFault(week, comp):
    drvrFault = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where ({REDACTED} = '0012' or (({REDACTED} = '058010' or {REDACTED} = '058010') and {REDACTED} = '10'))
		and {REDACTED} not like 'W%'
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    drvrFault = cursor.fetchone()

    if drvrFault[0] is None:
        return 0
    else:
        return drvrFault[0]

# to get towing costs not accident-related for one company slice for one week (repair reason 0051 or WAC 0020 excluding accident ROs)
def getNonAccTowing(week, comp):
    tow = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where ({REDACTED} = '0020' or {REDACTED} = '0051' or ({REDACTED} = '058010' or {REDACTED} = '058010'))
		and not ({REDACTED} in ('0031', '0032', '0036', '0037', '0039', '0012')
        or {REDACTED} = '0021'
        or ({REDACTED} = '000023' or {REDACTED} = '000023' or {REDACTED} = '000025' or {REDACTED} = '000025')
        or (({REDACTED} = '017000' or {REDACTED} = '017000') and {REDACTED} = '71'))
		and {REDACTED} not like 'W%'
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    tow = cursor.fetchone()

    if tow[0] is None:
        return 0
    else:
        return tow[0]

# to get pm costs for everything or just specifically tractors for one company slice for one week
def getPMCost(week, comp, tractor):
    pmCost = 0

    if tractor:
        sql = """
            select sum({REDACTED}) 
            from {REDACTED}
            join {REDACTED} on {REDACTED} = date
            join {REDACTED} on {REDACTED} = {REDACTED}
            where {REDACTED} = 'TRACTOR'
            and ({REDACTED} = '000002' or {REDACTED} = '000002' or {REDACTED} = '0008')
            and {REDACTED} not like 'W%'
            and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} """ + str(comp) + """
        """

    else:
        sql = """
            select sum({REDACTED}) 
            from {REDACTED}
            join {REDACTED} on {REDACTED} = date
            where ({REDACTED} = '000010' or {REDACTED} = '000010' or {REDACTED} = '000002' or {REDACTED} = '000002' or {REDACTED} = '000003' or {REDACTED} = '000003' or {REDACTED} = '000008' or {REDACTED} = '000008' or {REDACTED} = '0008')
            and {REDACTED} not like 'W%'
            and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} """ + str(comp) + """
        """

    cursor.execute(sql)

    pmCost = cursor.fetchone()

    if pmCost[0] is None:
        return 0
    else:
        return pmCost[0]

# to get tripac repair costs for one company slice for one week ({REDACTED} 000007, 082020, 082080)
def getTripac(week, comp):
    tripac = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where ({REDACTED} = '000007' or {REDACTED} = '000007' or {REDACTED} = '082020' or {REDACTED} = '082020' or {REDACTED} = '082080' or {REDACTED} = '082080')
		and {REDACTED} not like 'W%'
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    tripac = cursor.fetchone()

    if tripac[0] is None:
        return 0
    else:
        return tripac[0]

# to get brake repair costs for one company slice for one week ({REDACTED} starting with 013)
def getBrakes(week, comp):
    brakes = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where ({REDACTED} like '013%' or {REDACTED} like '013%')
		and {REDACTED} not like 'W%'
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    brakes = cursor.fetchone()

    if brakes[0] is None:
        return 0
    else:
        return brakes[0]

# to get power plant repair costs for one company slice for one week ({REDACTED} starting with 045)
def getPowerPlant(week, comp):
    pPlant = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where ({REDACTED} like '045%' or {REDACTED} like '045%')
		and {REDACTED} not like 'W%'
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    pPlant = cursor.fetchone()

    if pPlant[0] is None:
        return 0
    else:
        return pPlant[0]

# to get exhaust repair costs for one company slice for one week ({REDACTED} starting with 043)
def getExhaust(week, comp):
    exhaust = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where ({REDACTED} like '043%' or {REDACTED} like '043%')
		and {REDACTED} not like 'W%'
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    exhaust = cursor.fetchone()

    if exhaust[0] is None:
        return 0
    else:
        return exhaust[0]

# to get cumulative numbers for how mechanics spent their time according to timeclock data for one company slice for one week
def getProductivity(week, comp):
    prod = []

    sql = """
        select
        concat(ddd.year, '-', ddd.WeekOfYear) 'Week',
        sum(i.direct) 'Direct',
        sum(i.indirect) 'Indirect',
        sum(i.breaks) 'Breaks',
        sum(i.total) 'Total',
        sum(i.idle) 'Idle'

        from {REDACTED} ddd

        join
            (select 
            d.date,
            tc.{REDACTED},
            tc.{REDACTED},
            case when e.direct is null then 0 else e.direct end 'direct',
            case when f.indirect is null then 0 else f.indirect end 'indirect',
            case when h.breaks is null then 0 else h.breaks end 'breaks',
            case when sum(distinct g.total) is null then 0 else sum(distinct g.total) end 'total',
            case when
                sum(distinct g.total) - (case when e.direct is null then 0 else e.direct end + case when f.indirect is null then 0 else f.indirect end + case when h.breaks is null then 0 else h.breaks end)
                is null then 0 else
                sum(distinct g.total) - (case when e.direct is null then 0 else e.direct end + case when f.indirect is null then 0 else f.indirect end + case when h.breaks is null then 0 else h.breaks end)
                end 'idle'

            from
                {REDACTED} d
                join {REDACTED} tc on tc.{REDACTED} = d.date
                left join
                    (select	
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'direct'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on mm.{REDACTED} = tmt.{REDACTED}
                    and {REDACTED} = 'P'
                    and mm.{REDACTED} """ + str(comp) + """ 
                    and {REDACTED} = {REDACTED}
                    and {REDACTED} = 'SW'
                    group by date, tmt.{REDACTED}, {REDACTED}) e on e.date = d.date and e.{REDACTED} = tc.{REDACTED} and e.{REDACTED} = tc.{REDACTED}

                left join
                    (select	
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'indirect'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on tmt.{REDACTED} = mm.{REDACTED} 
                    and mm.{REDACTED} """ + str(comp) + """ 
                    and {REDACTED} = 'P'
                    and {REDACTED} = {REDACTED}
                    and {REDACTED} = 'ID'
                    group by date, tmt.{REDACTED}, {REDACTED}) f on f.date = d.date and f.{REDACTED} = tc.{REDACTED} and f.{REDACTED} = tc.{REDACTED}

                left join
                    (select	
                        date,
                        {REDACTED}, 
                        out.{REDACTED},
                        cast(((select 
                            max({REDACTED})
                        from {REDACTED} tme
                        join {REDACTED} da on {REDACTED} = date
                        where {REDACTED} = dd.date
                        and {REDACTED} = out.{REDACTED}
                        and tme.{REDACTED} = out.{REDACTED}
                        and {REDACTED} = {REDACTED}
                        and {REDACTED} = 'P'
                        and {REDACTED} = 'LO'
                        group by date)
                        -
                        (select
                            min({REDACTED})
                        from {REDACTED} tme
                        join {REDACTED} da on {REDACTED} = date
                        where {REDACTED} = dd.date
                        and {REDACTED} = {REDACTED}
                        and {REDACTED} = 'P'
                        and {REDACTED} = out.{REDACTED}
                        and tme.{REDACTED} = out.{REDACTED}
                        and {REDACTED} = 'LI'
                        group by date))
                        as float)
                        * 24 'total'

                    from {REDACTED} out
                    join {REDACTED} dd on out.{REDACTED} = dd.date
                    join {REDACTED} mm on mm.{REDACTED} = out.{REDACTED}
                    where mm.{REDACTED} """ + str(comp) + """ 
                    group by dd.date, {REDACTED}, out.{REDACTED}) g on g.date = d.date and g.{REDACTED} = tc.{REDACTED} and g.{REDACTED} = tc.{REDACTED}

                left join
                    (select
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'breaks'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on mm.{REDACTED} = tmt.{REDACTED}
                    and {REDACTED} = 'P'
                    and {REDACTED} = {REDACTED}
                    and mm.{REDACTED} """ + str(comp) + """ 
                    and {REDACTED} = 'LB'
                    group by date, tmt.{REDACTED}, {REDACTED}) h on h.date = d.date and h.{REDACTED} = tc.{REDACTED} and h.{REDACTED} = tc.{REDACTED}

                group by d.date, e.direct, f.indirect, h.breaks, tc.{REDACTED}, tc.{REDACTED}) i on ddd.date = i.Date

        where concat(ddd.year, ddd.WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        group by ddd.year, ddd.WeekOfYear
        order by ddd.year desc, ddd.WeekOfYear desc
    """

    cursor.execute(sql)

    prodData = cursor.fetchone()

    for item in prodData:
        prod.append(item)

    return prod

# to get either outside labor or outside part costs for one company slice for one week ({REDACTED} not 0001, type LB or PT)
def getOROCosts(week, comp, {REDACTED}):
    cost = 0

    sql = """
        select sum({REDACTED}) 
        from {REDACTED}
        join {REDACTED} on {REDACTED} = date
		where {REDACTED} != '0001'
		and {REDACTED} = '""" + str({REDACTED}) + """'
        and {REDACTED} not like 'W%'
        and not (({REDACTED} like '017%' or {REDACTED} like '017%') and {REDACTED} = 'PT')
        and not ({REDACTED} in ('0031', '0032', '0036', '0037', '0039', '0012')
        or {REDACTED} = '0021'
        or ({REDACTED} = '000023' or {REDACTED} = '000023' or {REDACTED} = '000025' or {REDACTED} = '000025')
        or (({REDACTED} = '017000' or {REDACTED} = '017000') and {REDACTED} = '71'))
		and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
		and {REDACTED} """ + str(comp) + """
    """

    cursor.execute(sql)

    cost = cursor.fetchone()

    if cost[0] is None:
        return 0
    else:
        return cost[0]

# to get true day miles ran by units for one company slice for one week
def getMiles(week, comp):
    miles = 0

    if comp == "not in ('31', '32')":
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            join {REDACTED} on ({REDACTED} = {REDACTED} and date = {REDACTED})
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} not in ('MTC', 'ENGT', 'ENM', 'ENMX')
        """

    elif comp == "= '1'":
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            join {REDACTED} on ({REDACTED} = {REDACTED} and date = {REDACTED})
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} = 'RKE'
        """

    elif comp == "= '11'":
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            join {REDACTED} on ({REDACTED} = {REDACTED} and date = {REDACTED})
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} = 'COL'
        """

    elif comp == "= '5'":
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            join {REDACTED} on ({REDACTED} = {REDACTED} and date = {REDACTED})
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} = 'ALB'
        """

    elif comp == "in ('3', '333')":
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            join {REDACTED} on ({REDACTED} = {REDACTED} and date = {REDACTED})
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} = 'EDE'
        """

    elif comp == "in ('31', '32')":
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            join {REDACTED} on ({REDACTED} = {REDACTED} and date = {REDACTED})
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
            and {REDACTED} in ('MTC', 'ENGT', 'ENM', 'ENMX')
        """
    
    else:
        sql = """
            select 
            sum([{REDACTED}])
            FROM {REDACTED} dm
            where concat(year, week) = '""" + str(week[0]) + str(week[1]) + """'
        """

    cursor.execute(sql)

    miles = cursor.fetchone()

    if miles[0] is None:
        return 0
    else:
        return miles[0]

# to get an array containing submitted and received warrany information for one company slice for one week (type starting with 'W' and {REDACTED} is 'RA' or 'SB')
def getWarrantyCosts(week, comp):
    warranty = []

    sql = """
        select concat(Year, '-', WeekOfYear) 'Week',
        (select 
            case
                when sum({REDACTED}) is null
            then 0
            else
                sum({REDACTED})
            end 'total_cost'
        from {REDACTED} 
        join {REDACTED} dds on {REDACTED} = date
        where {REDACTED} like 'W%'
        and {REDACTED} > 0
        and dds.Year = dd.Year
        and {REDACTED} """ + str(comp) + """
        and dds.WeekOfYear = dd.WeekOfYear
        and {REDACTED} in ('RA', 'SB')) 'total_cost',

        (select
            case
                when sum({REDACTED}) is null
            then 0
            else
                sum({REDACTED})
            end 'total_cost'
        from {REDACTED} 
        join {REDACTED} dds on {REDACTED} = date
        where {REDACTED} like 'W%'
        and {REDACTED} < 0
        and dds.Year = dd.Year
        and {REDACTED} """ + str(comp) + """
        and dds.WeekOfYear = dd.WeekOfYear
        and {REDACTED} in ('RA', 'SB')) * -1 'total_recieved'

        from {REDACTED} dd
        where concat(dd.year, dd.WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        group by dd.year, dd.WeekOfYear
        order by dd.year desc, dd.WeekOfYear desc
    """

    cursor.execute(sql)

    warrantyData = cursor.fetchone()

    for item in warrantyData:
        warranty.append(item)

    return warranty

# to get over the road and other costs and returns them as a 2 item list (vehicles in {REDACTED} 100, 600, and 601 vs all vehicles not in those {REDACTED})
def getOverTheRoadAndOtherCosts(week):
    overRoad = []

    sql = """
        select 
            (select
                sum({REDACTED}) 'overTheRoad'
                from {REDACTED} 
                join {REDACTED} on {REDACTED} = date
                join {REDACTED} on {REDACTED} = {REDACTED}
                where {REDACTED} in ('000100', '000600', '000601', '000610')
                and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
                and {REDACTED} not like 'W%') 'Over the Road',

            (select
                sum({REDACTED}) 'other'
                from {REDACTED} 
                join {REDACTED} on {REDACTED} = date
                join {REDACTED} on {REDACTED} = {REDACTED}
                where {REDACTED} not in ('000100', '000600', '000601', '000610')
                and concat(year, WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
                and {REDACTED} not like 'W%') 'Other'
    """

    cursor.execute(sql)

    overRoad = cursor.fetchone()
    
    return overRoad

# this function gets all vehicle years represented in the past years' data
def getYears(comp):
    years = []

    if comp == '102':
        company = "not in ('31', '32')"
    elif comp == '301':
        company = "in ('31', '32')"
    elif comp == 'RKE':
        company = "= '1'"
    elif comp == 'COL':
        company = "= '11'"
    elif comp == 'ALB':
        company = "= '5'"
    elif comp == 'EDE':
        company = "in ('3', '333')" 
    elif comp == 'ALL':
        company = "is not null"

    else:
        company = "is not null"

    sql = """
        declare @startDate date, @endDate date
            set @endDate = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

            set @startDate = '""" + str(lastYear) + """'

        select

        distinct {REDACTED} 'unit year'

        from {REDACTED} md

        join {REDACTED} on {REDACTED} = date

        join {REDACTED} on {REDACTED} = {REDACTED}

        where date between @startDate and @endDate

        and {REDACTED} """ + str(company) + """

        order by {REDACTED}
    """

    cursor.execute(sql)

    results = cursor.fetchall()

    for item in results:
        years.append(int(item[0]))

    return years

# to pull all indirect charge data for the last year for indirect charge page
def populateIndirectChargeData():
    idChargeData = []

    sql = """
        SELECT 
            concat(year, '-', WeekOfYear) 'Week',
            {REDACTED},
            ic.date, 
            {REDACTED},
            concat({REDACTED}, ' - ', {REDACTED}) 'WAC',
            {REDACTED},
            {REDACTED}
        FROM {REDACTED} ic
        join {REDACTED} d on d.date = ic.date
        join {REDACTED} on {REDACTED} = {REDACTED} and {REDACTED} = 109
        where ic.date between '""" + lastYear + """' and cast(dateadd(day, -2, '""" + str(startDate) + """') as date)
        order by ic.date desc, {REDACTED}
    """

    cursor.execute(sql)

    detailData = cursor.fetchall()

    for row in detailData:
        data = [row[0], row[1], row[2], row[3], row[4], row[5], row[6]]
        idChargeData.append(data)

    return idChargeData

# to pull details for all ROs for the last year for RO details page
def populateRoDetailData():
    roDetailData = []

    sql = """
        declare @startDate date, @endDate date
            set @endDate = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

            set @startDate = '""" + str(lastYear) + """'

        select
        concat(year, '-', WeekOfYear) 'week',
        {REDACTED},

        case 
            when {REDACTED} != '0001' then 'Y' else 'N' 
        end 'ORO',

        {REDACTED},
        {REDACTED},
        md.{REDACTED},
        {REDACTED},
        {REDACTED},
        {REDACTED} 'unit year',
        {REDACTED} 'equipment type',
        concat({REDACTED}, ' - ', v.{REDACTED}) 'Body System',
        concat({REDACTED}, ' - ', r.{REDACTED}) 'Repair Reason',
        {REDACTED},
        {REDACTED},
        {REDACTED},
        {REDACTED},
        {REDACTED},

        case 
            when {REDACTED} != '' then concat({REDACTED}, ' - ', w.{REDACTED}) else '' 
        end 'WAC',

        concat({REDACTED}, ' - ', l.{REDACTED}) 'Line Sys',

        case 
            when ({REDACTED} = {REDACTED} and {REDACTED} != '') then {REDACTED} when {REDACTED} = '' then '' when {REDACTED} = '002000999' then concat({REDACTED}, ' - ', 'UNKNOWN') 
            else case when (len({REDACTED}) = 6 and substring({REDACTED}, 4, 3) = '000') then 
            concat(substring({REDACTED}, 1, 3), ' - ', vmrdsc) else concat({REDACTED}, ' - ', vmrdsc) 
        end end 'VMRS Sys',

        {REDACTED},
        {REDACTED},

        case 
            when {REDACTED} = '0031' or {REDACTED} = '0032' then 'A'
            when {REDACTED} != '0001' and substring({REDACTED}, 1, 3) != '017' then 'O0'
            when {REDACTED} != '0001' and substring({REDACTED}, 1, 3) = '017' then 'TLO'
            when {REDACTED} = '0001' and substring({REDACTED}, 1, 3) != '017' then 'RO'
            when {REDACTED} = '0001' and substring({REDACTED}, 1, 3) = '017' then 'TLR'
        end 'Code'

        from {REDACTED} md

        join {REDACTED} on {REDACTED} = date

        join {REDACTED} r on r.{REDACTED} = {REDACTED}

        left join {REDACTED} w on w.{REDACTED} = {REDACTED} and (w.{REDACTED} = 15 or {REDACTED} = '')

        join {REDACTED} v on {REDACTED} = v.{REDACTED} and (v.{REDACTED} = 31)

        join {REDACTED} l on {REDACTED} = l.{REDACTED} and (l.{REDACTED} = 31)

        left join {REDACTED} on case when (len({REDACTED}) = 6 and substring({REDACTED}, 4, 3) = '000') then substring({REDACTED}, 1, 3) else {REDACTED} end = {REDACTED}

        left join {REDACTED} on {REDACTED} = {REDACTED}

        where date between @startDate and @endDate

        order by year desc, WeekOfYear desc, {REDACTED} desc
    """

    cursor.execute(sql)

    detailData = cursor.fetchall()

    for row in detailData:
        data = [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17],
                row[18], row[19], row[20], row[21], row[22]]
        roDetailData.append(data)

    return roDetailData

# to create a data array containing key data for key page
def getKeyData():
    keyData = []

    keyData.append(['Wk', 'The business week for which this row of data was aggregated'])
    keyData.append(['Accid', 'Sum of RO costs with {REDACTED} 0031, 0032, 0036, 0037, 0039, or 0012 or have {REDACTED} or {REDACTED} codes 000023 or 000025 or have {REDACTED} or {REDACTED} system code 000017 and {REDACTED} code 71'])
    keyData.append(['Tires', "Sum of RO costs with {REDACTED} PT and a {REDACTED} or {REDACTED} system code starting with '017'"])
    keyData.append(['ORO', "Sum of RO costs with {REDACTED} 0004, excluding accident ROs (detailed above) and tire ROs (also detailed above)"])
    keyData.append(['Maint & Repair', "Sum of RO costs which aren't accident, tire, or outside ROs (all detailed above)"])
    keyData.append(['Indirect', "Sum of non-RO labor charges (excluding tax and miscellaneous charges)"])
    keyData.append(['Total', "Sum of all 'COSTS' columns for the week"])
    keyData.append(['Years', "Sum of RO costs for units with that manufacturing year"])
    keyData.append(['Drvr Fault', "Sum of RO costs with {REDACTED} 0012 or with {REDACTED} or {REDACTED} system code 058010 and {REDACTED} code 10"])
    keyData.append(['Non-acc Tow', "Sum of RO costs excluding all accident ROs (detailed above) and with either {REDACTED} 0051 or {REDACTED} or {REDACTED} system code 058010"])
    keyData.append(['PM Trk', "Sum of RO costs where the vehicle type is 'TRACTOR' and either the {REDACTED} or {REDACTED} system code is 000002 or {REDACTED} is 0008"])
    keyData.append(['PM Total', "Sum of RO costs where either the {REDACTED} or {REDACTED} system code is 000002, 000003, 000008, or 000010, or {REDACTED} is 0008"])
    keyData.append(['Tripac', "Sum of RO costs where the {REDACTED} or {REDACTED} system code is 000007, 082020, or 082080"])
    keyData.append(['Brakes', "Sum of RO costs where the {REDACTED} or {REDACTED} system code begins with '013'"])
    keyData.append(['Power Plant', "Sum of RO costs where the {REDACTED} or {REDACTED} system code begins with '045'"])
    keyData.append(['Exhaust', "Sum of RO costs where the {REDACTED} or {REDACTED} system code begins with '043'"])
    keyData.append(['Dir Labor %', "Ratio of this week's direct labor hours to the sum of this week's direct and indirect labor hours"])
    keyData.append(['Unassign Labor Hrs', "Total clocked hours for the week minus direct hours, indirect hours, and breaks"])
    keyData.append(['Dir Labor Hrs', "Total hours spent assigned to ROs"])
    keyData.append(['Ind Labor Hrs', "Total hours spent assigned to indirect labor activities"])
    keyData.append(['ORO Parts', "The 'ORO' number but only including ROs with {REDACTED} PT"])
    keyData.append(['ORO Labor', "The 'ORO' number but only including ROs with {REDACTED} LB"])
    keyData.append(['Warranty Cost', "Sum of RO costs with {REDACTED} that begins with the letter 'W', {REDACTED} of SB or RA, and {REDACTED} greater than 0"])
    keyData.append(['Warranty Received', "Sum of RO costs with {REDACTED} that begins with the letter 'W', {REDACTED} of SB or RA and {REDACTED} less than 0, multiplied by -1"])
    keyData.append(['OTR Total', "Sum of RO costs for units under {REDACTED} 000100, 000600, and 000601"])
    keyData.append(['Non-OTR Total', "Sum of RO costs for units not under {REDACTED} 000100, 000600, and 000601"])
    keyData.append(['Miles', "Sum of true day miles run by units this week under this company/terminal"])
    keyData.append(['Cont CPM', "Total week's cost from above minus sum of accident ROs (detailed above), divided by the sum of true day miles"])
    keyData.append(['Total CPM', "Total week's cost divided by the sum of true day miles"])

    return keyData

# creates a 54-row array containing column headers followed by maintenance costs broken down by given week and company slice for the last year
def populateCompData(comp, weeks):
    company = ''
    data = []
    years = []

    # to pass into each SQL statement to only draw appropriate company slice
    if comp == '{REDACTED}':
        company = "not in ('31', '32')"
    elif comp == '{REDACTED}':
        company = "in ('31', '32')"
    elif comp == '{REDACTED}':
        company = "= '1'"
    elif comp == '{REDACTED}':
        company = "= '11'"
    elif comp == '{REDACTED}':
        company = "= '5'"
    elif comp == '{REDACTED}':
        company = "in ('3', '333')" 
    elif comp == '{REDACTED}':
        company = "is not null"
    else:
        # so this function still works for indirect charge tab
        data = populateIndirectChargeData()
        return data

    # gets manufacturing years represented in ROs for this particular company slice
    years = getYears(comp)

    ### build header row ###
    data.append(['Wk', 'Accid', 'Tires', 'ORO', 'Maint & Repair', 'Indirect', 'Total'])

    for year in years:
        data[0].append(str(year))
    
    data[0].append('Drvr Fault') 
    data[0].append('Non-acc Tow')
    data[0].append('PM Trk') 
    data[0].append('PM Total') 
    data[0].append('Tripac') 
    data[0].append('Brakes')  
    data[0].append('Power Plant')
    data[0].append('Exhaust') 
    data[0].append('Dir Labor %')
    data[0].append('Unassign Labor Hrs') 
    data[0].append('Dir Labor Hrs') 
    data[0].append('Ind Labor Hrs')
    data[0].append('ORO Parts') 
    data[0].append('ORO Labor') 
    data[0].append('Warranty Cost') 
    data[0].append('Warranty Recieved') 

    # only for overview page
    if comp == 'ALL':
        data[0].append('OTR Total')
        data[0].append('Non-OTR Total')

    data[0].append('Miles')
    data[0].append('Cont CPM')
    data[0].append('Total CPM')

    # this variable is used to indicate index of list so data pulled is written correctly to list at the end of this method
    increment = 1

    # for every week in this data, starting with last week:
    for week in weeks:
        warranty = []
        mfgYears = []
        productivity = []

        weekNum = str(week[0]) + '-' + str(week[1])
        accid = getAcc(week, company)
        tires = getTires(week, company)
        oro = getORO(week, company)
        mAndR = getMntRep(week, company)
        indirect = getIndirectCosts(week, company)

        # these data points are aggregated from data already pulled instead of doing more SQL
        totalNonAcc = tires + oro + mAndR + indirect
        total = accid + tires + oro + mAndR + indirect

        # adds data for each year to a list to be inserted into this row of data at the end of this method
        for year in years:
            mfgYears.append(getMFGYear(week, year, company))

        drvrFault = getDriverFault(week, company)
        nonAccTow = getNonAccTowing(week, company)
        trkPMs = getPMCost(week, company, True)
        allPMs = getPMCost(week, company, False)
        tripak = getTripac(week, company)
        brakes = getBrakes(week, company)
        power = getPowerPlant(week, company)
        exhaust = getExhaust(week, company)

        productivity = getProductivity(week, company)

        # catches potential divide by zero error
        if productivity[1] != 0:
            dirLabPct = productivity[1] / (productivity[1] + productivity[2])

        else:
            dirLabPct = 0

        unassignLabHrs = productivity[5]
        dirLab = productivity[1]
        indLab = productivity[2]
        oroPts = getOROCosts(week, company, 'PT')
        oroLb = getOROCosts(week, company, 'LB')
        
        warranty = getWarrantyCosts(week, company)

        warrantySub = warranty[1]
        warrantyRec = warranty[2]

        overVsOther = getOverTheRoadAndOtherCosts(week)

        overRoad = overVsOther[0]
        other = overVsOther[1]

        miles = getMiles(week, company)

        # catches potential divide by zero error
        if totalNonAcc != 0:
            contCPM = round((totalNonAcc / miles), 3)
        else:
            contCPM = 0

        # catches potential divide by zero error
        if total != 0:
            totalCPM = round((total / miles), 3)
        else:
            totalCPM = 0

        # assemble all this week's gathered data into a row
        data.append([weekNum, accid, tires, oro, mAndR, indirect, total])

        for year in mfgYears:
            data[increment].append(year)

        data[increment].append(drvrFault) 
        data[increment].append(nonAccTow) 
        data[increment].append(trkPMs) 
        data[increment].append(allPMs)
        data[increment].append(tripak) 
        data[increment].append(brakes) 
        data[increment].append(power) 
        data[increment].append(exhaust) 
        data[increment].append(dirLabPct) 
        data[increment].append(unassignLabHrs) 
        data[increment].append(dirLab) 
        data[increment].append(indLab) 
        data[increment].append(oroPts) 
        data[increment].append(oroLb) 
        data[increment].append(warrantySub) 
        data[increment].append(warrantyRec) 

        # only for overview page
        if comp == 'ALL':
            data[increment].append(overRoad)
            data[increment].append(other)

        data[increment].append(miles) 
        data[increment].append(contCPM) 
        data[increment].append(totalCPM)

        increment += 1

    return data

#### end data collecting methods ####

#### begin construction methods ####

# takes given worksheet and data array and makes correctly-formatted header row out of the top row of the data
def buildHeaderRow(worksheet, data):
    col = 0

    for item in data[0]:
        if item == 'Wk':
            worksheet.write('A5', str(item), header_fmt_center_white)
        else:
            worksheet.write(4, col, str(item), header_fmt_center)

        col += 1

# takes given worksheet and data array and sets column widths, hides mfg year and incident columns, and sets the background format
def setColWidths(sheet, data):
    col = 0

    for item in data[0]:
        if not (str(item[0]) == '1' or str(item[0]) == '2' or str(item) == 'Drvr Fault' or str(item) == 'Non-acc Tow' or str(item) == 'Non-warranty Credits'):
            sheet.set_column(col, col, 12, background_fmt)

        else:
            sheet.set_column(col, col, 12, background_fmt, {'hidden': 1})

        col += 1

    sheet.set_column('A:BZ', None, background_fmt)

# creates array of dictionaries containing proper total functions and formats for table-building method of xlsxwriter based on first row of given data array
def getTotalFunctions(data):
    columns = []

    for item in data[0]:
        if str(item) == 'Wk':
            columns.append({})  

        elif str(item) == 'Dir Labor %':
            columns.append({'total_function': 'average',
                            'format': percent_fmt})

        elif str(item) == 'Miles':
            columns.append({'total_function': 'sum',
                            'format': reg_fmt})

        elif str(item) == 'Cont CPM' or str(item) == 'Total CPM':
            columns.append({'total_function': 'average',
                            'format': dec_fmt})

        elif str(item) == 'Unassign Labor Hrs' or str(item) == 'Dir Labor Hrs' or str(item) == 'Ind Labor Hrs':
            columns.append({'total_function': 'sum',
                            'format': hour_fmt})
        else:
            columns.append({'total_function': 'sum',
                            'format': money_fmt})

    return columns

# writes and correctly formats data in data array to given worksheet
def writeData(sheet, data, comp):
    inc = 0
    rowInd = 5
    years = getYears(comp)

    for row in data:
        # skip the first row of data because the first row contains the column headers
        if inc == 0:
            inc += 1
            continue
        
        col = 0
        sheet.write(rowInd, col, row[col], everything_else_border)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt_border)
        col += 1

        for year in years:
            if year != years[-1]:
                sheet.write(rowInd, col, row[col], money_fmt)
                col += 1

            else:
                sheet.write(rowInd, col, row[col], money_fmt_border)
                col += 1

        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt_border)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt_border)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt_border)
        col += 1
        sheet.write(rowInd, col, row[col], percent_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], hour_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], hour_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], hour_fmt_border)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt_border)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], money_fmt_border)
        col += 1

        # only for overview page
        if comp == 'ALL':
            sheet.write(rowInd, col, row[col], money_fmt)
            col += 1
            sheet.write(rowInd, col, row[col], money_fmt_border)
            col += 1

        sheet.write(rowInd, col, row[col], reg_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], dec_fmt)
        col += 1
        sheet.write(rowInd, col, row[col], dec_fmt_border)
        col += 1
        
        rowInd += 1

    sheet.write('A59', 'Totals:', header_fmt_center_white)

# uses the number of columns in given data array and given company string to build a title row at the top of each sheet
def writeTitles(comp, sheet, data):

    if comp == '102':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Fleetmaster', title_fmt)

    elif comp == '301':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Englander', title_fmt)

    elif comp == 'RKE':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Roanoke', title_fmt)

    elif comp == 'COL':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Columbus', title_fmt)

    elif comp == 'ALB':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Albany', title_fmt)

    elif comp == 'EDE':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Eden', title_fmt)

    elif comp == 'ALL':
        sheet.merge_range(2, 1, 2, len(data[0]) - 1, 'Maintenance KPI - Overview', title_fmt)

# uses given data array to build and correctly format a subheader row to improve visual organization
def writeSubHeaders(sheet, data):
    costs = 'B4:G4'
    OTRSt = 0
    OTREd = 0
    mfgYearSt = 7
    mfgYearEd = 6

    for item in data[0]:
        if str(item[0]) == '1' or str(item[0]) == '2':
            mfgYearEd += 1

    incidentsSt = mfgYearEd + 1
    incidentsEd = incidentsSt + 1

    pmsSt = incidentsEd + 1
    pmsEd = pmsSt + 1

    systemsSt = pmsEd + 1
    systemsEd = systemsSt + 3

    productivitySt = systemsEd + 1
    productivityEd = productivitySt + 3

    outVendorSt = productivityEd + 1
    outVendorEd = outVendorSt + 1

    warrantySt = outVendorEd + 1
    warrantyEd = warrantySt + 1

    # only for overview page
    if data[0][-5] == 'OTR Total':
        OTRSt = warrantyEd + 1
        OTREd = OTRSt + 1

        cpmSt = OTREd + 1
        cpmEd = cpmSt + 2

    else:
        cpmSt = warrantyEd + 1
        cpmEd = cpmSt + 2

    sheet.merge_range(costs, 'COSTS', subheader_fmt)
    sheet.merge_range(3, mfgYearSt, 3, mfgYearEd, 'MFG YEAR', subheader_fmt)
    sheet.merge_range(3, incidentsSt, 3, incidentsEd, 'INCIDENTS', subheader_fmt)
    sheet.merge_range(3, pmsSt, 3, pmsEd, 'PMs', subheader_fmt)
    sheet.merge_range(3, systemsSt, 3, systemsEd, 'SYSTEMS', subheader_fmt)
    sheet.merge_range(3, productivitySt, 3, productivityEd, 'PRODUCTIVITY', subheader_fmt)
    sheet.merge_range(3, outVendorSt, 3, outVendorEd, 'OUTSIDE VENDOR', subheader_fmt)
    sheet.merge_range(3, warrantySt, 3, warrantyEd, 'WARRANTY', subheader_fmt)

    # only for overview page
    if OTRSt != 0:
        sheet.merge_range(3, OTRSt, 3, OTREd, 'OTR', subheader_fmt)

    sheet.merge_range(3, cpmSt, 3, cpmEd, 'CPM', subheader_fmt)

# def addImage(sheet):
#     if sheet is engSheet:
#         #sheet.insert_image('Y1', 'C:/{REDACTED}',{'x_scale': .13, 'y_scale': .13}) # for production
#         sheet.insert_image('Y1', 'O:/{REDACTED}',{'x_scale': .13, 'y_scale': .13}) # for local testing
#     else:
#         #sheet.insert_image('Y1', 'C:/{REDACTED}',{'x_scale': 0.2, 'y_scale': 0.2}) # for production
#         sheet.insert_image('Y1', 'O:/{REDACTED}',{'x_scale': 0.2, 'y_scale': 0.2}) # for local testing
        
# builds table using given data to determine subheaders and total functions on given worksheet
def buildTable(sheet, data):
    sheet.freeze_panes(5, 1)
    # addImage(sheet)

    buildHeaderRow(sheet, data)
    columns = getTotalFunctions(data)

    sheet.set_row(2, 40)
    sheet.set_row(3, 30)

    sheet.add_table(5, 0, 58, len(data[0]) - 1, {'header_row': False,
                                                'style': 'Table Style Light 1',
                                                'total_row': True,
                                                'columns': columns})

#### end construction methods ####

#### begin chart building methods ####

# get a list of all mechanic codes represented in past years' timeclock data
def getMechanics():
    mechanics = []
    sql = """
        declare @startDate date, @endDate date
            set @endDate = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

            set @startDate = '""" + str(lastYear) + """'

        select distinct tc.{REDACTED}
        from {REDACTED} tc
        join {REDACTED} m on tc.{REDACTED} = m.{REDACTED}

        where {REDACTED} between @startDate and @endDate
        and {REDACTED} = 'A'
    """

    cursor.execute(sql)

    mechData = cursor.fetchall()

    for row in mechData:
        mechanics.append(str(row[0]))

    return mechanics

# get year's worth of hours for given mechanic
def getMechanicAll(mechanic):
    breakdown = []

    sql = """
        declare @startDate date, @endDate date
            set @endDate = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

            set @startDate = '""" + str(lastYear) + """'

        select
        concat(ddd.year, '-', ddd.WeekOfYear) 'Week',
		concat(i.{REDACTED}, ' - ', mm.{REDACTED}) 'Mechanic',
        sum(i.direct) 'Direct',
        sum(i.indirect) 'Indirect',
        sum(i.breaks) 'Breaks',
        sum(i.total) 'Total',
        sum(i.idle) 'Idle'

        from {REDACTED} ddd

        join
            (select 
            d.date,
            tc.{REDACTED},
            tc.{REDACTED},
            case when e.direct is null then 0 else e.direct end 'direct',
            case when f.indirect is null then 0 else f.indirect end 'indirect',
            case when h.breaks is null then 0 else h.breaks end 'breaks',
            case when sum(distinct g.total) is null then 0 else sum(distinct g.total) end 'total',
            case when
                sum(distinct g.total) - (case when e.direct is null then 0 else e.direct end + case when f.indirect is null then 0 else f.indirect end + case when h.breaks is null then 0 else h.breaks end)
                is null then 0 else
                sum(distinct g.total) - (case when e.direct is null then 0 else e.direct end + case when f.indirect is null then 0 else f.indirect end + case when h.breaks is null then 0 else h.breaks end)
                end 'idle'

            from
                {REDACTED} d
                join {REDACTED} tc on tc.{REDACTED} = d.date
                left join
                    (select	
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'direct'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on mm.{REDACTED} = tmt.{REDACTED}
                    and {REDACTED} = 'P'
                    and mm.{REDACTED} is not null
                    and {REDACTED} = {REDACTED}
                    and {REDACTED} = 'SW'
                    group by date, tmt.{REDACTED}, {REDACTED}) e on e.date = d.date and e.{REDACTED} = tc.{REDACTED} and e.{REDACTED} = tc.{REDACTED}

                left join
                    (select	
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'indirect'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on tmt.{REDACTED} = mm.{REDACTED} 
                    and mm.{REDACTED} is not null
                    and {REDACTED} = 'P'
                    and {REDACTED} = {REDACTED}
                    and work_s{REDACTED}tatus = 'ID'
                    group by date, tmt.{REDACTED}, {REDACTED}) f on f.date = d.date and f.{REDACTED} = tc.{REDACTED} and f.{REDACTED} = tc.{REDACTED}

                left join
                    (select	
                        date,
                        {REDACTED}, 
                        out.{REDACTED},
                        cast(((select 
                            max({REDACTED})
                        from {REDACTED} tme
                        join {REDACTED} da on {REDACTED} = date
                        where {REDACTED} = dd.date
                        and {REDACTED} = out.{REDACTED}
                        and tme.{REDACTED} = out.{REDACTED}
                        and {REDACTED} = {REDACTED}
                        and {REDACTED} = 'P'
                        and {REDACTED} = 'LO'
                        group by date)
                        -
                        (select
                            min({REDACTED})
                        from {REDACTED} tme
                        join {REDACTED} da on {REDACTED} = date
                        where {REDACTED} = dd.date
                        and {REDACTED} = {REDACTED}
                        and {REDACTED} = 'P'
                        and {REDACTED} = out.{REDACTED}
                        and tme.{REDACTED} = out.{REDACTED}
                        and {REDACTED} = 'LI'
                        group by date))
                        as float)
                        * 24 'total'

                    from {REDACTED} out
                    join {REDACTED} dd on out.{REDACTED} = dd.date
                    join {REDACTED} mm on mm.{REDACTED} = out.{REDACTED}
                    where mm.{REDACTED} is not null
                    group by dd.date, {REDACTED}, out.{REDACTED}) g on g.date = d.date and g.{REDACTED} = tc.{REDACTED} and g.{REDACTED} = tc.{REDACTED}

                left join
                    (select
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'breaks'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on mm.{REDACTED} = tmt.{REDACTED}
                    and {REDACTED} = 'P'
                    and {REDACTED} = {REDACTED}
                    and mm.{REDACTED} is not null
                    and {REDACTED} = 'LB'
                    group by date, tmt.{REDACTED}, {REDACTED}) h on h.date = d.date and h.{REDACTED} = tc.{REDACTED} and h.{REDACTED} = tc.{REDACTED}

                group by d.date, e.direct, f.indirect, h.breaks, tc.{REDACTED}, tc.{REDACTED}) i on ddd.date = i.Date

				join {REDACTED} mm on i.{REDACTED} = mm.{REDACTED}

        where ddd.date between @startDate and @endDate
        and i.{REDACTED} = '""" + mechanic + """'
        group by ddd.year, ddd.WeekOfYear, i.{REDACTED}, mm.{REDACTED}
        order by ddd.year, ddd.WeekOfYear, i.{REDACTED}
    """

    cursor.execute(sql)

    mechBreakdown = cursor.fetchall()

    for row in mechBreakdown:
        breakdown.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6]])

    return breakdown

# get all mechanics' hour breadowns for a single week 
def getHourBreakdown(week):
    breakdown = []

    sql = """
        declare @startDate date, @endDate date
            set @endDate = cast(dateadd(day, -2, '""" + str(startDate) + """') as date)

            set @startDate = '""" + str(lastYear) + """'

        select
        concat(ddd.year, '-', ddd.WeekOfYear) 'Week',
		concat(i.{REDACTED}, ' - ', mm.{REDACTED}) 'Mechanic',
        sum(i.direct) 'Direct',
        sum(i.indirect) 'Indirect',
        sum(i.breaks) 'Breaks',
        sum(i.total) 'Total',
        sum(i.idle) 'Idle'

        from Dim.DateDimension ddd

        join
            (select 
            d.date,
            tc.{REDACTED},
            tc.{REDACTED},
            case when e.direct is null then 0 else e.direct end 'direct',
            case when f.indirect is null then 0 else f.indirect end 'indirect',
            case when h.breaks is null then 0 else h.breaks end 'breaks',
            case when sum(distinct g.total) is null then 0 else sum(distinct g.total) end 'total',
            case when
                sum(distinct g.total) - (case when e.direct is null then 0 else e.direct end + case when f.indirect is null then 0 else f.indirect end + case when h.breaks is null then 0 else h.breaks end)
                is null then 0 else
                sum(distinct g.total) - (case when e.direct is null then 0 else e.direct end + case when f.indirect is null then 0 else f.indirect end + case when h.breaks is null then 0 else h.breaks end)
                end 'idle'

            from
                {REDACTED} d
                join {REDACTED} tc on tc.{REDACTED} = d.date
                left join
                    (select	
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'direct'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on mm.{REDACTED} = tmt.{REDACTED}
                    and {REDACTED} = 'P'
                    and mm.{REDACTED} is not null
                    and {REDACTED} = {REDACTED}
                    and {REDACTED} = 'SW'
                    group by date, tmt.{REDACTED}, {REDACTED}) e on e.date = d.date and e.{REDACTED} = tc.{REDACTED} and e.{REDACTED} = tc.{REDACTED}

                left join
                    (select	
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'indirect'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on tmt.{REDACTED} = mm.{REDACTED} 
                    and mm.{REDACTED} is not null
                    and {REDACTED} = 'P'
                    and {REDACTED} = {REDACTED}
                    and {REDACTED} = 'ID'
                    group by date, tmt.{REDACTED}, {REDACTED}) f on f.date = d.date and f.{REDACTED} = tc.{REDACTED} and f.{REDACTED} = tc.{REDACTED}

                left join
                    (select	
                        date,
                        {REDACTED}, 
                        out.{REDACTED},
                        cast(((select 
                            max({REDACTED})
                        from {REDACTED} tme
                        join {REDACTED} da on {REDACTED} = date
                        where {REDACTED} = dd.date
                        and {REDACTED} = out.{REDACTED}
                        and tme.{REDACTED} = out.{REDACTED}
                        and {REDACTED} = {REDACTED}
                        and {REDACTED} = 'P'
                        and {REDACTED} = 'LO'
                        group by date)
                        -
                        (select
                            min({REDACTED})
                        from {REDACTED} tme
                        join {REDACTED} da on {REDACTED} = date
                        where {REDACTED} = dd.date
                        and {REDACTED} = {REDACTED}
                        and {REDACTED} = 'P'
                        and {REDACTED} = out.{REDACTED}
                        and tme.{REDACTED} = out.{REDACTED}
                        and {REDACTED} = 'LI'
                        group by date))
                        as float)
                        * 24 'total'

                    from {REDACTED} out
                    join {REDACTED} dd on out.{REDACTED} = dd.date
                    join {REDACTED} mm on mm.{REDACTED} = out.{REDACTED}
                    where mm.{REDACTED} is not null
                    group by dd.date, {REDACTED}, out.{REDACTED}) g on g.date = d.date and g.{REDACTED} = tc.{REDACTED} and g.{REDACTED} = tc.{REDACTED}

                left join
                    (select
                        date,
                        tmt.{REDACTED},
                        {REDACTED},
                        sum(cast(({REDACTED} - {REDACTED}) as float) * 24) 'breaks'

                    from {REDACTED} tmt
                    join {REDACTED} on {REDACTED} = date
                    join {REDACTED} mm on mm.{REDACTED} = tmt.{REDACTED}
                    and {REDACTED} = 'P'
                    and {REDACTED} = {REDACTED}
                    and mm.{REDACTED} is not null
                    and {REDACTED} = 'LB'
                    group by date, tmt.{REDACTED}, {REDACTED}) h on h.date = d.date and h.{REDACTED} = tc.{REDACTED} and h.{REDACTED} = tc.{REDACTED}

                group by d.date, e.direct, f.indirect, h.breaks, tc.{REDACTED}, tc.{REDACTED}) i on ddd.date = i.Date

				join {REDACTED} mm on i.{REDACTED} = mm.{REDACTED}

        where concat(ddd.year, ddd.WeekOfYear) = '""" + str(week[0]) + str(week[1]) + """'
        group by ddd.year, ddd.WeekOfYear, i.{REDACTED}, mm.{REDACTED}
        order by ddd.year, ddd.WeekOfYear, i.{REDACTED}
    """

    cursor.execute(sql)

    weekBreakdown = cursor.fetchall()

    for row in weekBreakdown:
        breakdown.append([row[0], row[1].lower().title(), row[2], row[3], row[4], row[5], row[6]])

    return breakdown

# Creates active mechanic breakdown chart
def create_bar_chart_worksheet(workbook, worksheet, name, start_position, bg_format, data):
    # Create a new chart object.
    chart = workbook.add_chart({'type': 'column'})

    chart.add_series({'name': 'Direct', 'categories': '=ChartData!$K$2:$K$' + str(len(data) + 1), 'values': '=ChartData!$L$2$:$L$' + str(len(data) + 1), 'gap': 200, 'overlap': -40})
    chart.add_series({'name': 'Indirect', 'categories': '=ChartData!$K$2:$K$' + str(len(data) + 1), 'values': '=ChartData!$M$2$:$M$' + str(len(data) + 1)})
    chart.add_series({'name': 'Breaks', 'categories': '=ChartData!$K$2:$K$' + str(len(data) + 1), 'values': '=ChartData!$N$2$:$N$' + str(len(data) + 1)})
    chart.add_series({'name': 'Total', 'categories': '=ChartData!$K$2:$K$' + str(len(data) + 1), 'values': '=ChartData!$O$2$:$O$' + str(len(data) + 1)})
    chart.add_series({'name': 'Idle', 'categories': '=ChartData!$K$2:$K$' + str(len(data) + 1), 'values': '=ChartData!$P$2$:$P$' + str(len(data) + 1)})

    chart.set_title({
    'name': name,
    'name_font': {
        'name': 'Calibri',
        'color': 'white'
        },
    })
    chart.set_plotarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_chartarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_x_axis({'num_font': {'color': 'white', 'size': 9}, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_legend({'font': {'name': 'Calibri', 'color': 'white', 'size': 9}})
    chart.set_y_axis({'name': 'Hours', 'num_font': {'color': 'white'}, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    
    # changing x and y scale here changes the actual size of the chart
    worksheet.insert_chart(start_position, chart, {'x_scale': 2.5, 'y_scale': 1.3})

    worksheet.set_column('A:AF', None, bg_format)
    return workbook

# create individual mechanic hour breakdown charts
def create_line_chart_worksheet(workbook, worksheet, name, beginning, end, start_position, bg_format):
    # Create a new chart object.
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'name': 'Direct', 'categories': '=ChartData!$A$' + str(beginning) + ':$A$' + str(end), 'values': '=ChartData!$C$' + str(beginning) + ':$C$' + str(end), 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Indirect', 'categories': '=ChartData!$A$' + str(beginning) + ':$A$' + str(end), 'values': '=ChartData!$D$' + str(beginning) + ':$D$' + str(end), 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Breaks', 'categories': '=ChartData!$A$' + str(beginning) + ':$A$' + str(end), 'values': '=ChartData!$E$' + str(beginning) + ':$E$' + str(end), 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Total', 'categories': '=ChartData!$A$' + str(beginning) + ':$A$' + str(end), 'values': '=ChartData!$F$' + str(beginning) + ':$F$' + str(end), 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Idle', 'categories': '=ChartData!$A$' + str(beginning) + ':$A$' + str(end), 'values': '=ChartData!$G$' + str(beginning) + ':$G$' + str(end), 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    
    chart.set_title({
    'name': name,
    'name_font': {
        'name': 'Calibri',
        'color': 'white'
        },
    })
    chart.set_plotarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_chartarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_x_axis({'num_font': {'color': 'white', 'rotation': -90, 'size': 9}, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_y_axis({'name': 'Hours', 'num_font': {'color': 'white'}, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_legend({'font': {'name': 'Calibri', 'color': 'white', 'size': 9}})
    
    # changing x and y scale here changes the actual size of the chart
    worksheet.insert_chart(start_position, chart, {'x_scale': 2.5, 'y_scale': 1.3})

    worksheet.set_column('A:AF', None, bg_format)
    return workbook

# create Total Maintenance Cost Shop Comparison chart
def create_total_chart_worksheet(workbook, worksheet, start_position, bg_format):
    # Create a new chart object.
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'name': 'Overview', 'categories': '=Overview!$A$6:$A$58', 'values': '=Overview!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Fleetmaster', 'categories': '=Overview!$A$6:$A$58', 'values': '=Fleetmaster!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Englander', 'categories': '=Overview!$A$6:$A$58', 'values': '=Englander!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Roanoke', 'categories': '=Overview!$A$6:$A$58', 'values': '=Roanoke!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Columbus', 'categories': '=Overview!$A$6:$A$58', 'values': '=Columbus!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Albany', 'categories': '=Overview!$A$6:$A$58', 'values': '=Albany!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Eden', 'categories': '=Overview!$A$6:$A$58', 'values': '=Eden!$G$6:$G$58', 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    
    chart.set_title({
    'name': 'Total Maintenance Cost Shop Comparison',
    'name_font': {
        'name': 'Calibri',
        'color': 'white'
        },
    })
    chart.set_plotarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_chartarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_x_axis({'num_font': {'color': 'white', 'rotation': -90, 'size': 9}, 'reverse': True, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_y_axis({'num_font': {'color': 'white'}, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_legend({'font': {'name': 'Calibri', 'color': 'white', 'size': 9}})
    
    # changing x and y scale here changes the actual size of the chart
    worksheet.insert_chart(start_position, chart, {'x_scale': 2.5, 'y_scale': 1.3})

    worksheet.set_column('A:AF', None, bg_format)
    return workbook

# create CPM comparison charts
def create_CPM_chart_worksheet(workbook, worksheet, columnCont, columnTot, name, start_position, bg_format):
    # Create a new chart object.
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'name': 'Cont', 'categories': '=Overview!$A$6:$A$58', 'values': ['Overview', 5, columnCont, 57, columnCont], 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    chart.add_series({'name': 'Total', 'categories': '=Overview!$A$6:$A$58', 'values': ['Overview', 5, columnTot, 57, columnTot], 'marker': {'type': 'circle', 'fill': {'color': 'white'}, 'border': {'color': 'black'}}})
    
    chart.set_title({
    'name': name,
    'name_font': {
        'name': 'Calibri',
        'color': 'white'
        },
    })
    chart.set_plotarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_chartarea({'gradient': {'colors': ['black', 'black']}})
    chart.set_x_axis({'num_font': {'color': 'white', 'rotation': -90, 'size': 9}, 'reverse': True, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_y_axis({'name': 'Dollar Amount per Mile', 'num_font': {'color': 'white'}, 'name_font': {'name': 'Calibri', 'color': 'white'}})
    chart.set_legend({'font': {'name': 'Calibri', 'color': 'white', 'size': 9}})
    
    # changing x and y scale here changes the actual size of the chart
    worksheet.insert_chart(start_position, chart, {'x_scale': 2.5, 'y_scale': 1.3})

    worksheet.set_column('A:AF', None, bg_format)
    return workbook

# gets index of Cont CPM column
def getContInd(data):
    ind = 0

    for row in data[0]:
        if row != 'Cont CPM':
            ind += 1

    return ind - 1

# get index of Total CPM column
def getTotInd(data):
    ind = 0

    ind = len(data[0]) - 1

    return ind

#### end chart building methods ####

### end methods ###

### all formats for this report ###

# centered column header format
header_fmt_center = writer.add_format({ 
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#DF1B16',
    'font_color': '#FFFFFF',
    'font_name': 'Sans-Serif',
    'text_wrap': True,
    'font_size': 10,
    'top': 0})

# format for total cell at bottom left of tables
header_fmt_center_white = writer.add_format({ 
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FFFFFF',
    'font_color': '#000000',
    'font_name': 'Sans-Serif',
    'text_wrap': True,
    'font_size': 10})

# format for header categories
subheader_fmt = writer.add_format({ 
    'bold': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#DF1B16',
    'font_color': '#FFFFFF',
    'font_name': 'Sans-Serif',
    'text_wrap': True,
    'font_size': 13,
    'border': 1,
    'bottom': 0})

# title format
title_fmt = writer.add_format({ 
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#000000',
    'font_color': '#FFFFFF',
    'font_size': 25})

# format for chart page headers
chart_header_fmt = writer.add_format({
    'bold': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#262626',
    'font_color': '#FFFFFF',
    'font_size': 18,
    'underline': 1
})

# format for headers on indirect charge, RO Detail, and ChartData pages
header_fmt = writer.add_format({'bold': 1,
                                'fg_color': '#DF1B16',
                                'font_color': '#FFFFFF',
                                'font_name': 'Sans-Serif',
                                'text_wrap': True,
                                'font_size': 10})

# format to apply text-wrapping to cell
wrap_fmt = writer.add_format({'text_wrap': True})

# format for chart backgrounds
background_fmt = writer.add_format({'bg_color': '#262626',
                                    'font_size': 10})

# format to apply to columns with percentage values
percent_fmt = writer.add_format({'num_format': '#0.00%',
                                 'font_size': 10,
                                 'align': 'right'}) 

# format to apply to columns whose number value may go above 1,000
reg_fmt = writer.add_format({'num_format': '#,###,##0.0',
                             'font_size': 10,
                             'align': 'right'}) 

# format for CPM columns
dec_fmt = writer.add_format({'num_format': '0.#00',
                             'font_size': 10})

# format for CPM column on far right
dec_fmt_border = writer.add_format({'num_format': '0.#00',
                                    'font_size': 10,
                                    'right': 1})

# format for productivity columns
hour_fmt = writer.add_format({'num_format': '#,##0.00',
                              'font_size': 10,
                              'align': 'right'})

# format for far-right productivity column
hour_fmt_border = writer.add_format({'num_format': '#,##0.00',
                                     'font_size': 10,
                                     'align': 'right',
                                     'right': 1})


# format to apply to columns that contain currency values
money_fmt = writer.add_format({'num_format': '[$$-409]#,##0.00',
                               'font_size': 10,
                               'align': 'right'}) 

# format to apply to column that contains currency values on far right of subheader section
money_fmt_border = writer.add_format({'num_format': '[$$-409]#,##0.00',
                                      'font_size': 10,
                                      'align': 'right',
                                      'right': 1})

# format to center column values
center_fmt = writer.add_format({'align': 'center',
                                'font_size': 10}) 

# format for date columns
date_fmt = writer.add_format({'num_format': 'yyyy-mm-dd',
                              'align': 'center',
                              'font_size': 10})

# format to change font size for otherwise unformatted columns
everything_else = writer.add_format({'font_size': 10,
                                     'align': 'right',
                                     'bottom': 0})

# format to change font size for otherwise unformatted columns on far right of subheader section
everything_else_border = writer.add_format({'font_size': 10,
                                            'align': 'right',
                                            'right': 1,
                                            'bottom': 0})

### end formats ###

### begin main area ###

### assemble data lists for report ###

try:
    print('Getting RO Detail Data...')
    roDetailData = populateRoDetailData()
    print('Done.')
    print('Getting Overview Data...')
    allData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Fleetmaster Data...')
    fltData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Englander Data...')
    engData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Roanoke terminal Data...')
    rkeData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Columbus terminal Data...')
    colData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Albany terminal Data...')
    albData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Eden terminal Data...')
    edeData = populateCompData('{REDACTED}', weeks)
    print('Done.')
    print('Getting Indirect Charge Data...')
    indData = populateCompData('{REDACTED}', weeks)
    print('Done.')

    ### end data assembly/collection ###

    ### create sheets of report ###
    print('Building worksheets and tables...')
    ro_detail_sheet = writer.add_worksheet('RO Details')
    allSheet = writer.add_worksheet('Overview')
    fltSheet = writer.add_worksheet('Fleetmaster')
    engSheet = writer.add_worksheet('Englander')
    rkeSheet = writer.add_worksheet('Roanoke')
    colSheet = writer.add_worksheet('Columbus')
    albSheet = writer.add_worksheet('Albany')
    edeSheet = writer.add_worksheet('Eden')
    indSheet = writer.add_worksheet('Indirect Labor')
    chartSheet = writer.add_worksheet('Charts')
    chartData = writer.add_worksheet('ChartData')
    keySheet = writer.add_worksheet('Key')

    ### end sheet creation ###

    ### prepare shop sheets for table data

    writeTitles('{REDACTED}', allSheet, allData)
    writeTitles('{REDACTED}', fltSheet, fltData)
    writeTitles('{REDACTED}', engSheet, engData)
    writeTitles('{REDACTED}', rkeSheet, rkeData)
    writeTitles('{REDACTED}', colSheet, colData)
    writeTitles('{REDACTED}', albSheet, albData)
    writeTitles('{REDACTED}', edeSheet, edeData)

    setColWidths(allSheet, allData)
    setColWidths(fltSheet, fltData)
    setColWidths(engSheet, engData)
    setColWidths(rkeSheet, rkeData)
    setColWidths(colSheet, colData)
    setColWidths(albSheet, albData)
    setColWidths(edeSheet, edeData)

    writeSubHeaders(allSheet, allData)
    writeSubHeaders(fltSheet, fltData)
    writeSubHeaders(engSheet, engData)
    writeSubHeaders(rkeSheet, rkeData)
    writeSubHeaders(colSheet, colData)
    writeSubHeaders(albSheet, albData)
    writeSubHeaders(edeSheet, edeData)

    ### end shop sheet preparation ###

    ### build non-shop sheets and tables ###

    ro_detail_sheet.set_column('A:A', 9)
    ro_detail_sheet.set_column('B:B', 9)
    ro_detail_sheet.set_column('C:C', 9)
    ro_detail_sheet.set_column('D:D', 8)
    ro_detail_sheet.set_column('E:E', 8)
    ro_detail_sheet.set_column('F:F', 11)
    ro_detail_sheet.set_column('G:G', 15)
    ro_detail_sheet.set_column('H:H', 8)
    ro_detail_sheet.set_column('I:I', 13)
    ro_detail_sheet.set_column('J:J', 15)
    ro_detail_sheet.set_column('K:K', 39)
    ro_detail_sheet.set_column('L:L', 31)
    ro_detail_sheet.set_column('M:M', 10)
    ro_detail_sheet.set_column('N:N', 9)
    ro_detail_sheet.set_column('O:O', 19)
    ro_detail_sheet.set_column('P:P', 7)
    ro_detail_sheet.set_column('Q:Q', 9)
    ro_detail_sheet.set_column('R:R', 31)
    ro_detail_sheet.set_column('S:S', 39)
    ro_detail_sheet.set_column('T:T', 76)
    ro_detail_sheet.set_column('U:U', 10)
    ro_detail_sheet.set_column('V:V', 11)
    ro_detail_sheet.set_column('W:W', 8)

    ro_detail_sheet.freeze_panes(1, 0)

    ro_detail_sheet.add_table('A1:W' + str(len(roDetailData) + 1), {'data': roDetailData,
                                                                    'style': 'Table Style Light 1',
                                                                    'columns': [{'header': 'Week:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'Shop:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'ORO?',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'RO#:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Sect:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'RO Sts:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'Comp Date:',
                                                                                'header_format': header_fmt,
                                                                                'format': date_fmt},
                                                                                {'header': 'Unit:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Mfg. Year:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'Equip Type:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'Body System:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Repair Reason:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Charge Date:',
                                                                                'header_format': header_fmt,
                                                                                'format': date_fmt},
                                                                                {'header': 'Type:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt},
                                                                                {'header': 'Mech/Part:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Qty:',
                                                                                'header_format': header_fmt,
                                                                                'format': reg_fmt},
                                                                                {'header': 'Cost:',
                                                                                'header_format': header_fmt,
                                                                                'format': money_fmt},
                                                                                {'header': 'WAC:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Line Sys:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'VMRS Sys:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'RO ID#:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Line ID#:',
                                                                                'header_format': header_fmt,
                                                                                'format': everything_else},
                                                                                {'header': 'Code:',
                                                                                'header_format': header_fmt,
                                                                                'format': center_fmt}]})

    indSheet.set_column('A:A', 8)
    indSheet.set_column('B:B', 8)
    indSheet.set_column('C:C', 10)
    indSheet.set_column('D:D', 12)
    indSheet.set_column('E:E', 36)
    indSheet.set_column('F:F', 9)
    indSheet.set_column('G:G', 9)

    indSheet.freeze_panes(1, 0)

    indSheet.add_table('A1:G' + str(len(indData) + 1), {'data': indData,
                                                        'style': 'Table Style Light 1',
                                                        'columns': [{'header': 'Week:',
                                                                    'header_format': header_fmt,
                                                                    'format': center_fmt},
                                                                    {'header': 'Shop:',
                                                                    'header_format': header_fmt},
                                                                    {'header': 'Date:',
                                                                    'header_format': header_fmt,
                                                                    'format': date_fmt},
                                                                    {'header': 'Mechanic:',
                                                                    'header_format': header_fmt,
                                                                    'format': center_fmt},
                                                                    {'header': 'WAC:',
                                                                    'header_format': header_fmt},
                                                                    {'header': 'Hours:',
                                                                    'header_format': header_fmt},
                                                                    {'header': 'Cost:',
                                                                    'header_format': header_fmt,
                                                                    'format': money_fmt}]})

    keyData = getKeyData()

    keySheet.set_column('A:A', 25)
    keySheet.set_column('B:B', 110)

    keySheet.freeze_panes(1, 0)

    keySheet.add_table('A1:B' + str(len(keyData) + 1), {'data': keyData,
                                                        'style': 'Table Style Light 1',
                                                        'autofilter': False,
                                                        'columns': [{'header': 'Column:',
                                                                    'header_format': header_fmt},
                                                                    {'header': 'Description:',
                                                                    'header_format': header_fmt,
                                                                    'format': wrap_fmt}]})

    ### end non-shop table building ###

    ### finally build shop sheet tables and write data into them and do final non-chart-based processing ###

    buildTable(allSheet, allData)
    buildTable(fltSheet, fltData)
    buildTable(engSheet, engData)
    buildTable(rkeSheet, rkeData)
    buildTable(albSheet, albData)
    buildTable(colSheet, colData)
    buildTable(edeSheet, edeData)

    writeData(allSheet, allData, '{REDACTED}')
    writeData(fltSheet, fltData, '{REDACTED}')
    writeData(engSheet, engData, '{REDACTED}')
    writeData(rkeSheet, rkeData, '{REDACTED}')
    writeData(colSheet, colData, '{REDACTED}')
    writeData(albSheet, albData, '{REDACTED}')
    writeData(edeSheet, edeData, '{REDACTED}')

    rkeSheet.hide()
    colSheet.hide()
    albSheet.hide()
    edeSheet.hide()
    indSheet.hide()
    chartData.hide()

    ### end final shop page processing ###

    print('Done.') 
    print('Writing ChartData and building charts...')

    ### start chart building activities ###

    # returns list of all active mechanics
    mechanics = getMechanics()

    # gets list of hour breakdowns for all mechanics for last week
    hourBreakdown = getHourBreakdown(weeks[0])

    # will contain all individual mechanic hour breakdowns for every week in past year
    mechanicData = []

    ### To break down what this next section of code is doing: it takes every mechanic in list we got above and tries to get an hour breakdown for every business week 
    ### for the past year. Problem is, not all active mechanics have time posted in every week for the past year - the SQL will only return data for weeks for which 
    ### that mechanic has time posted, and not zeroes or nulls for those weeks either - it'll just return nothing. Given that we need 53 weeks of data for each chart,
    ### we have to fill the blanks with the right business week and 5 zeroes for the chart building function to make the charts correctly. So, this takes the dataset 
    ### returned by getMechanicALL() and checks the first index of each row (which contains the business week), comparing it to what that week year should be if it's
    ### correctly ordered (using the already-properly-ordered weeks list) and if that business week isn't in the data, it inserts a row with that business week and 
    ### five zeroes, which is accurate since they did no work that week - otherwise, it just inserts the data for that week that the function returned. It decrements 
    ### the weekNum variable after adding each row into mechanicData. It also checks at the end of returned data that weekNum has reached zero - if it hasn't by the
    ### end of the returned data, it adds rows with every remaining business week and five zeroes and decrements weekNum each time until mechanicData has a row for
    ### every week for the past year. In other words, by the time weekNum is 0, it has ensured that there is a row of data for each of the last 52 weeks for every
    ### active mechanic and written it into the mechanicData list to be used to create the table on the ChartData page.
    ### Thanks for attending my Ted talk.
    
    for mechanic in mechanics:
        mechanicHourData = getMechanicAll(mechanic)
        weekNum = 52

        for row in mechanicHourData:

            while row[0] != str(weeks[weekNum][0]) + '-' + str(weeks[weekNum][1]):
                mechanicData.append([str(weeks[weekNum][0]) + '-' + str(weeks[weekNum][1]), row[1], 0.0, 0.0, 0.0, 0.0, 0.0])
                weekNum -= 1

            mechanicData.append(row)
            weekNum -= 1
            if weekNum != 0 and row == mechanicHourData[-1]:

                while weekNum > -1:
                    mechanicData.append([str(weeks[weekNum][0]) + '-' + str(weeks[weekNum][1]), row[1], 0.0, 0.0, 0.0, 0.0, 0.0])
                    weekNum -= 1

    # builds table using list assembled directly above, just in case the ChartData page is accidentally unhidden
    chartData.add_table('A1:G' + str(len(mechanicData) + 1), {'data': mechanicData,
                                                            'style': 'Table Style Light 1',
                                                            'autofilter': False,
                                                            'columns': [{'header': 'Week',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Mechanic',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Direct',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Indirect',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Breaks',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Total',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Idle',
                                                                        'header_format': header_fmt}]})

    # builds table for data to build single week hour breakdown chart in case ChartData page is accidentally unhidden
    chartData.add_table('J1:P' + str(len(hourBreakdown) + 1), {'data': hourBreakdown,
                                                            'style': 'Table Style Light 1',
                                                            'autofilter': False,
                                                            'columns': [{'header': 'Week',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Mechanic',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Direct',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Indirect',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Breaks',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Total',
                                                                        'header_format': header_fmt},
                                                                        {'header': 'Idle',
                                                                        'header_format': header_fmt}]})

    # header for mechanic hours charts
    chartSheet.merge_range('L2:R3', 'Mechanic Hours Charts', chart_header_fmt)

    # creates top chart on chart sheet
    create_bar_chart_worksheet(writer, chartSheet, 'All Mechanic Hours Breakdown - Week ' + str(weeks[0][0]) + '-' + str(weeks[0][1]), 'F5', background_fmt, hourBreakdown)

    beg = 2
    end = 54
    chartLoc = 25

    # builds and adds chart for each mechanic and increments chart and data location markers
    for mechanic in mechanics:
        create_line_chart_worksheet(writer, chartSheet, mechanicData[beg][1].lower().title() + ' Hours', beg, end, 'F' + str(chartLoc), background_fmt)
        beg += 53
        end += 53
        chartLoc += 20

    chartLoc += 5

    # header for cost charts
    chartSheet.merge_range('L' + str(chartLoc - 3) + ':R' + str(chartLoc - 2), 'Maintenance Cost Charts', chart_header_fmt)

    # creates chart comparing shop totals for past year
    create_total_chart_worksheet(writer, chartSheet, 'F' + str(chartLoc), background_fmt)

    # keep incrementing chart marker so it places charts correctly on page
    chartLoc += 20

    # these hold indexes of CPM columns to build below chart comparing CPMs
    contCol = getContInd(allData)
    totCol = getTotInd(allData)

    # creates CPM comparison chart
    create_CPM_chart_worksheet(writer, chartSheet, contCol, totCol, 'Overview CPM Comparison', 'F' + str(chartLoc), background_fmt)

    ### end chart building activities ###

    print('Done.')

    # closing the workbook so it can be put into email and sent
    writer.close() 

    # subject of the email to be sent 
    eSubject = "Maintenance KPI for period ending with " + str(startDate - timedelta(days = 3)) 

    # body of email to be sent
    bodyHtml = """</table><br><br><br><br><i>{REDACTED} """ + str(st) + """<i>"""

    email = emailing("{REDACTED}", eSubject, bodyHtml, 'Maintenance_KPI_' + str(startDate) + '.xlsx', 'Maintenance_KPI_' + str(startDate) + '.xlsx')
    
    try:
        email.send_mail()
    except:
        worked = False
        while worked == False:
            try:
                email.send_mail
                worked = True
            except:
                pass
    finally:
        # delete created document to avoid clutter
        os.remove('Maintenance_KPI_' + str(startDate) + '.xlsx')

# if anything goes wrong:
except Exception as e:

    email = emailing("{REDACTED}", "WARNING - Maintenance KPI error notice", "<center><b>WARNING</b> the Maintenance KPI has encountered an error. <br><h1><b>Error:</b></h1>" + str(e))
    
    try:
        email.send_mail()
    except:
        worked = False
        while worked == False:
            try:
                email.send_mail
                worked = True
            except:
                pass
    finally:
        # delete created document to avoid clutter
        os.remove('Maintenance_KPI_' + str(startDate) + '.xlsx')
