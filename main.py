import requests
from lxml import etree
from icalendar import Calendar, Event
import datetime
from uuid import uuid1

s = requests.session()
header={
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36",
    'Content-Type': "application/x-www-form-urlencoded"
}
nbsp = "\xa0"

week2num = {
    "星期一":1,
    "星期二":2,
    "星期三":3,
    "星期四":4,
    "星期五":5,
    "星期六":6,
    "星期天":7,
}

# 根据变动设置每节课的时间
sect2datatime = {
    1:{
        "start_hour":8,
        "start_min":30,
        "end_hour":9,
        "end_min":15,
    },
    2:{
        "start_hour":9,
        "start_min":25,
        "end_hour":10,
        "end_min":10,
    },
    3:{
        "start_hour":10,
        "start_min":30,
        "end_hour":11,
        "end_min":15,
    },
     4:{
        "start_hour":11,
        "start_min":25,
        "end_hour":12,
        "end_min":10,
    },
     6:{
        "start_hour":14,
        "start_min":30,
        "end_hour":15,
        "end_min":15,
    },
     7:{
        "start_hour":15,
        "start_min":25,
        "end_hour":16,
        "end_min":10,
    },
    8:{
        "start_hour":16,
        "start_min":30,
        "end_hour":17,
        "end_min":15,
    },
    9:{
    "start_hour":17,
        "start_min":25,
        "end_hour":18,
        "end_min":10,
    },
    10:{
    "start_hour":19,
        "start_min":00,
        "end_hour":19,
        "end_min":45,
    },
    11:{
    "start_hour":19,
        "start_min":55,
        "end_hour":20,
        "end_min":40,
    },
    12:{
    "start_hour":20,
        "start_min":50,
        "end_hour":21,
        "end_min":35,
    }

}

#设置第一周第一天的时间
first_day = datetime.datetime(year=2023,month=2,day=20)

def getClassInfo(payload:dict,store:bool=False):
    """获取课程信息的html文件

    Args:
        payload (dict): post参数
        store (bool, optional): 是否保存得到的html
    Returns:
        _type_: 返回response
    """
    response = s.post(url="http://mis.cqu.edu.cn/mis/curricula/show_stu.jsp",data=payload,headers=header)
    if store:
        with open('classInfo.html','w',encoding='utf-8') as fp:
            fp.write(response.text)
        print("succesfully store classes infomation")
    return response

def parseFromHtml(html:str,path:str=None)->dict:
    """从html解析信息，可以直接传入，也可传入文件路径

    Args:
        html (str): 提供html
        path (str, optional): 也可以选择提供文件路径

    Returns:
        dict: 按照星期来排列课程的dict
    """
    week = {
        "星期一":[],
        "星期二":[],
        "星期三":[],
        "星期四":[],
        "星期五":[],
        "星期六":[],
        "星期天":[],
    }
    if path is not None:
        with open(path,'r',encoding='utf-8') as fp:
            content = fp.read()
    else:
        content = html
    root = etree.HTML(content)
    table = root.xpath('//table')[0]
    col = table.xpath('./tr[2]/td/text()')[1:]
    classes = table.xpath('./tr[position() >= 3]')
    for item in classes:
        for i in range(2,9):
            infos = item.xpath(f'./td[{i}]/text()')
            if infos[0]==nbsp:
                continue
            else:
                temp = {}
                for j in range(len(infos)):
                    info = infos[j]
                    [attr,cont] = info.strip().split('：')
                    # 遍历到最后插入
                    if j == len(infos)-1:
                        temp[attr.strip()] = cont.strip()
                        week[col[i-2]].append(temp)
                    # 碰到班号先把之前的插入，但是第一个不处理
                    elif "班号" in info and len(temp):
                        week[col[i-2]].append(temp)
                        temp = {}
                        temp[attr.strip()] = cont.strip()
                    else:
                        temp[attr.strip()] = cont.strip()
    return week

def sectAndWeek2Datatime(weekday,week,start_sect,end_sect):
    """根据周和节来计算时间

    Args:
        weekday (_type_): 星期几
        week (_type_): 最开始的一周
        start_sect (_type_): 开始节数
        end_sect (_type_): 结束的节数

    Returns:
        _type_: 返回开始时间和结束时间
    """
    start_week = (int(week)-1)*7 + (week2num[weekday]-1)
    start_sect_time = sect2datatime[int(start_sect)]
    end_sect_time = sect2datatime[int(end_sect)]
    day = first_day + datetime.timedelta(days=start_week)
    start_time = day + datetime.timedelta(hours=start_sect_time["start_hour"],minutes=start_sect_time["start_min"])
    end_time = day + datetime.timedelta(hours=end_sect_time["end_hour"],minutes=end_sect_time["end_min"])
    return start_time, end_time

def weekFliter(week):
    """对连堂课程进行合并

    Args:
        week (_type_): 传入week这个dict

    Returns:
        _type_: 返回去重后的week
    """
    res = {
        "星期一":[],
        "星期二":[],
        "星期三":[],
        "星期四":[],
        "星期五":[],
        "星期六":[],
        "星期天":[]
    }
    for key,value in week.items():
        temp = {}
        for cls in value:
            uid = cls["名称"]+cls["周次"]
            if temp.get(uid) is None:
                temp[uid] = cls
            else:
                temp[uid]["节次"] = temp[uid]["节次"][0]+'-'+cls["节次"][-1]
        res[key] = [v for k,v in temp.items()]
    return res
        
    
def generateICS(week):
    """根据week生成ICS

    Args:
        week (_type_): 包含课程信息的week dict
    """
    # 创建日历
    MyCalender = Calendar()
    # 添加属性
    MyCalender.add('X-WR-CALNAME', '课程表')
    MyCalender.add('prodid', '-//My calendar//luan//CN')
    MyCalender.add('version', '2.0')
    MyCalender.add('METHOD', 'PUBLISH')
    MyCalender.add('CALSCALE', 'GREGORIAN')  # 历法：公历
    # 遍历一天
    for key,value in week.items():
        # 遍历一天的课程
        for cls in value:
            cls_weeks = cls["周次"][:-1]
            cls_sect = cls["节次"]
            cls_weeks = cls_weeks.strip().split(" ")
            [start_sect,end_sect] = cls_sect.split("-")
            # 处理周数是1-2,4-6,12这种情况
            # 如果是12，那么开始和结束都是12
            for cls_week in cls_weeks:
                if cls_week.isdigit():
                    start_time, end_time = sectAndWeek2Datatime(key,cls_week,start_sect,end_sect)
                    weekend = cls_week
                    weekstart = cls_week
                else:
                    [weekstart,weekend] = cls_week.split("-")
                    start_time, end_time = sectAndWeek2Datatime(key,weekstart,start_sect,end_sect)
                # 添加事件
                event = Event()
                event_uuid = str(uuid1())+'@class'
                event.add('uid', event_uuid)
                event.add('summary', cls["名称"])
                event.add('dtstart', start_time)
                event.add('dtend', end_time)
                event.add('location', cls["教室"])
                event.add('description', '授课老师：'+cls['教师'])
                event.add('rrule', {'freq': 'weekly',
                                    'interval': 1,
                                    'count': int(weekend)-int(weekstart)+1})
                MyCalender.add_component(event)
    # print(calendar.get_ics_text())
    # 写入文件
    with open('课程表.ics', 'wb') as file:
        file.write(MyCalender.to_ical().replace(b'\r\n', b'\n').strip())
        print('导出完毕！')

def getStuSerial(id:str,password:str)->int:
    """根据学号和密码获取stuSerial来获取课表

    Args:
        id (str): 学号
        password (str): 密码

    Returns:
        int: stuSerial
    """
    url = "http://mis.cqu.edu.cn/mis/login.jsp"
    payload = {
        'userId':id,
        'password':password,
        'userType':'student',
    }
    s.post(url, data=payload, headers=header)
    res = s.get(url="http://mis.cqu.edu.cn/mis/student_manage.jsp")
    root = etree.HTML(res.text)
    x = root.xpath("//input[@value='我的课表']")[0]
    temp = x.xpath('string(./@onclick)')
    stuSerial = ''.join([i for i in temp if i.isdigit()])
    return stuSerial

def getICS(userId:int,password:str,term:int,week_distinct:bool=True):
    """生成ICS

    Args:
        userId (int): 学号
        password (str): 密码
        term (int): 学期，1 或者2
        week_distinct (bool, optional): 是否合并连堂课程，默认为True
    """
    stuSerial = getStuSerial(userId,password)
    payload = {
        'term': term,
        'week': None,
        'stuSerial': stuSerial
    }
    res = getClassInfo(payload=payload)
    week = parseFromHtml(html=res.text)
    if week_distinct:
        week = weekFliter(week)
    generateICS(week)



    
if __name__ == '__main__':
    userId = None
    password = None
    term = 2
    getICS(userId,password,term)
    




