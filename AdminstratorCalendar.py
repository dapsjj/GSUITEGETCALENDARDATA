from __future__ import print_function
import time
import logging
import os
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import re


def write_log():
    '''
    写log
    :return: 返回logger对象
    '''
    # 获取logger实例，如果参数为空则返回root logger
    logger = logging.getLogger()
    now_date = datetime.datetime.now().strftime('%Y%m%d')
    log_file = now_date+".log"# 文件日志
    if not os.path.exists("log"):#python文件同级别创建log文件夹
        os.makedirs("log")
    # 指定logger输出格式
    formatter = logging.Formatter('%(asctime)s %(levelname)s line:%(lineno)s %(message)s')
    file_handler = logging.FileHandler("log" + os.sep + log_file, mode='a', encoding='utf-8')
    file_handler.setFormatter(formatter) # 可以通过setFormatter指定输出格式
    # 为logger添加的日志处理器，可以自定义日志处理器让其输出到其他地方
    logger.addHandler(file_handler)
    # 指定日志的最低输出级别，默认为WARN级别
    logger.setLevel(logging.INFO)
    return logger

def removeBlank(MyString):
    '''
    :param MyString: 要替换空白的字符串
    :return: 去掉空白后的字符串
    '''
    MyString = re.sub('[\s+]', '', MyString)
    return MyString



def getcalendarIdList():
    '''
    读人员邮箱地址
    '''
    txt_config_list = []
    fileName = r"./CalendarIdList.txt"
    try:
        with open(fileName, 'r', encoding='utf-8') as txtConfig:
            lines = txtConfig.readlines()
            for line in lines:
                line = line.strip()
                if not line:  # 如果line是空
                    continue
                else:
                    txt_config_list.append(line)
            return txt_config_list
    except Exception as ex:
        logger.error("Call method getcalendarIdList() error!")
        raise ex


def generateEveryDayCalendarData(GmailList):
    '''
    :param GmailList: API中calendarId对应的List
    :return: 要保存成文本的List
    '''
    if GmailList:
        store = file.Storage('token.json')  # 如果换了证书的json文件，需要是删除token.json
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('./Administrator.json', SCOPES)
            creds = tools.run_flow(flow, store)
        service = build('calendar', 'v3', http=creds.authorize(Http()))
        TokyoStartTime = 'T00:00:00+09:00' #东京时间
        TokyoEndTime = 'T23:59:59+09:00' #东京时间
        UTCStartTime = 'T15:00:00.000Z' #UTC时间
        UTCEndTime = 'T14:59:59.999Z' #UTC时间

        List_Root_And_Branch = [] #Root节点和Branch节点拼接到一起
        # if not os.path.exists(CalendarPath + Yesterday):
        #     os.makedirs(CalendarPath + Yesterday)
        if not os.path.exists(CalendarPath):
            os.makedirs(CalendarPath)

        for calendarId in GmailList:
            # events_result = service.events().list(calendarId=calendarId,
            #                                     timeMin = Yesterday + StartTime,
            #                                     timeMax = Yesterday + EndTime,
            #                                     orderBy = 'updated', #orderBy允许的值['startTime', 'updated']
            #                                     showHiddenInvitations = True,
            #                                     showDeleted=True
            #                                     ).execute()

            events_result = service.events().list(calendarId=calendarId,
                                                  orderBy='updated',  # orderBy允许的值['startTime', 'updated']
                                                  showHiddenInvitations=True,
                                                  showDeleted=True,
                                                  updatedMin = Yesterday + TokyoStartTime
                                                  ).execute()
            RootNode_kind = removeBlank(str(events_result.get('kind','_'))) #dict.get(key, default=None)
            RootNode_etag = removeBlank(str(events_result.get('etag','_')))
            RootNode_summary = removeBlank(str(events_result.get('summary','_')))
            RootNode_description = removeBlank(str(events_result.get('description','_')))
            RootNode_updated = removeBlank(str(events_result.get('updated','_')))
            RootNode_timeZone = removeBlank(str(events_result.get('timeZone','_')))
            RootNode_accessRole = removeBlank(str(events_result.get('accessRole','_')))
            RootNode_defaultReminders = removeBlank(str(events_result.get('defaultReminders','[]')))#数组
            RootNode_nextPageToken = removeBlank(str(events_result.get('nextPageToken','_')))
            RootNode_nextSyncToken = removeBlank(str(events_result.get('nextSyncToken','_')))
            RootNode_items = events_result.get('items')#Root节点下的items节点,是个数组

            #items下的节点List
            List_BranchNode_items = []
            if RootNode_items: #判断items是否为空
                for item in RootNode_items:
                    EveryBranchNodeItem = []#每个items中的节点
                    Add_ele_to_BranchNodeItem = []#给EveryBranchNodeItem添加前面的元素用的List

                    BranchNode_items_kind = item.get('kind','_')
                    BranchNode_items_etag = item.get('etag','_')
                    BranchNode_items_id = item.get('id','_')
                    BranchNode_items_status = item.get('status','_')
                    BranchNode_items_htmlLink = item.get('htmlLink','_')
                    BranchNode_items_created = item.get('created','_')
                    BranchNode_items_updated = item.get('updated','_')
                    BranchNode_items_summary = item.get('summary','_')
                    BranchNode_items_description = item.get('description','_')
                    BranchNode_items_location = item.get('location','_')
                    BranchNode_items_colorId = item.get('colorId','_')
                    # BranchNode_items_creator = item.get('creator','_')
                    BranchNode_items_creator_id = item.get('creator',{}).get('id','_')
                    BranchNode_items_creator_email = item.get('creator',{}).get('email','_')
                    BranchNode_items_creator_displayName = item.get('creator',{}).get('displayName','_')
                    BranchNode_items_creator_self = item.get('creator',{}).get('self','_')
                    # BranchNode_items_organizer = item.get('organizer','_')
                    BranchNode_items_organizer_id = item.get('organizer', {}).get('id', '_')
                    BranchNode_items_organizer_email = item.get('organizer', {}).get('email', '_')
                    BranchNode_items_organizer_displayName = item.get('organizer', {}).get('displayName', '_')
                    BranchNode_items_organizer_self = item.get('organizer', {}).get('self', '_')
                    # BranchNode_items_start = item.get('start','_')
                    BranchNode_items_start_date = item.get('start', {}).get('date', '_')
                    BranchNode_items_start_dateTime = item.get('start', {}).get('dateTime', '_')
                    BranchNode_items_start_timeZone = item.get('start', {}).get('timeZone', '_')
                    # BranchNode_items_end = item.get('end','_')
                    BranchNode_items_end_date = item.get('end', {}).get('date', '_')
                    BranchNode_items_end_dateTime = item.get('end', {}).get('dateTime', '_')
                    BranchNode_items_end_timeZone = item.get('end', {}).get('timeZone', '_')
                    BranchNode_items_endTimeUnspecified = item.get('endTimeUnspecified','_')
                    BranchNode_items_recurrence = item.get('recurrence','[]')#数组
                    BranchNode_items_recurringEventId = item.get('recurringEventId','_')
                    BranchNode_items_originalStartTime_date = item.get('originalStartTime', {}).get('date', '_')
                    BranchNode_items_originalStartTime_dateTime = item.get('originalStartTime', {}).get('dateTime', '_')
                    BranchNode_items_originalStartTime_timeZone = item.get('originalStartTime', {}).get('timeZone', '_')
                    BranchNode_items_transparency = item.get('transparency','_')
                    BranchNode_items_visibility = item.get('visibility','_')
                    BranchNode_items_iCalUID = item.get('iCalUID','_')
                    BranchNode_items_sequence = item.get('sequence','_')
                    BranchNode_items_attendees = item.get('attendees','[]')
                    BranchNode_items_attendeesOmitted = item.get('attendeesOmitted','_')
                    # BranchNode_items_extendedProperties = item.get('extendedProperties','_')
                    BranchNode_items_extendedProperties_private = item.get('extendedProperties',{}).get('private','_')
                    BranchNode_items_extendedProperties_shared = item.get('extendedProperties',{}).get('shared','_')
                    BranchNode_items_hangoutLink = item.get('hangoutLink','_')
                    # BranchNode_items_conferenceData = item.get('conferenceData','_')
                    BranchNode_items_conferenceData_createRequest_requestId = item.get('conferenceData',{}).get('createRequest',{}).get('requestId','_')
                    BranchNode_items_conferenceData_createRequest_conferenceSolutionKey_type = item.get('conferenceData',{}).get('createRequest',{}).get('conferenceSolutionKey',{}).get('type','_')
                    BranchNode_items_conferenceData_createRequest_status_statusCode = item.get('conferenceData',{}).get('createRequest',{}).get('status',{}).get('statusCode','_')
                    BranchNode_items_conferenceData_entryPoints = item.get('conferenceData',{}).get('entryPoints','[]') #数组
                    # BranchNode_items_conferenceSolution = item.get('conferenceData',{}).get('conferenceSolution','_')
                    BranchNode_items_conferenceData_conferenceSolution_key_type = item.get('conferenceData',{}).get('conferenceSolution',{}).get('key',{}).get('type','_')
                    BranchNode_items_conferenceData_conferenceSolution_name = item.get('conferenceData',{}).get('conferenceSolution',{}).get('name','_')
                    BranchNode_items_conferenceData_conferenceSolution_iconUri = item.get('conferenceData',{}).get('conferenceSolution',{}).get('iconUri','_')
                    BranchNode_items_conferenceData_conferenceId = item.get('conferenceData',{}).get('conferenceId','_')
                    BranchNode_items_conferenceData_signature = item.get('conferenceData',{}).get('signature','_')
                    BranchNode_items_conferenceData_notes = item.get('conferenceData',{}).get('notes','_')
                    # BranchNode_items_gadget = item.get('gadget','_')
                    BranchNode_items_gadget_type = item.get('gadget',{}).get('type','_')
                    BranchNode_items_gadget_title = item.get('gadget',{}).get('title','_')
                    BranchNode_items_gadget_link = item.get('gadget',{}).get('link','_')
                    BranchNode_items_gadget_iconLink = item.get('gadget',{}).get('iconLink','_')
                    BranchNode_items_gadget_width = item.get('gadget',{}).get('width','_')
                    BranchNode_items_gadget_height = item.get('gadget',{}).get('height','_')
                    BranchNode_items_gadget_display = item.get('gadget',{}).get('display','_')
                    BranchNode_items_gadget_preferences = item.get('gadget',{}).get('preferences','_')
                    BranchNode_items_anyoneCanAddSelf = item.get('anyoneCanAddSelf','_')
                    BranchNode_items_guestsCanInviteOthers = item.get('guestsCanInviteOthers','_')
                    BranchNode_items_guestsCanModify = item.get('guestsCanModify','_')
                    BranchNode_items_guestsCanSeeOtherGuests = item.get('guestsCanSeeOtherGuests','_')
                    BranchNode_items_privateCopy = item.get('privateCopy','_')
                    BranchNode_items_locked = item.get('locked','_')
                    # BranchNode_reminders = item.get('reminders','_')
                    BranchNode_reminders_useDefault = item.get('reminders',{}).get('useDefault','_')
                    BranchNode_reminders_overrides = item.get('reminders',{}).get('overrides','[]')#数组
                    # BranchNode_source = item.get('source','_')
                    BranchNode_source_url = item.get('source',{}).get('url','_')
                    BranchNode_source_title = item.get('source',{}).get('title','_')
                    BranchNode_attachments = item.get('attachments','[]')#数组

                    #添加到分支节点的List中
                    EveryBranchNodeItem =[BranchNode_items_kind,
                                          BranchNode_items_etag,
                                          BranchNode_items_id,
                                          BranchNode_items_status,
                                          BranchNode_items_htmlLink,
                                          BranchNode_items_created,
                                          BranchNode_items_updated,
                                          BranchNode_items_summary,
                                          BranchNode_items_description,
                                          BranchNode_items_location,
                                          BranchNode_items_colorId,
                                          BranchNode_items_creator_id,
                                          BranchNode_items_creator_email,
                                          BranchNode_items_creator_displayName,
                                          BranchNode_items_creator_self,
                                          BranchNode_items_organizer_id,
                                          BranchNode_items_organizer_email,
                                          BranchNode_items_organizer_displayName,
                                          BranchNode_items_organizer_self,
                                          BranchNode_items_start_date,
                                          BranchNode_items_start_dateTime,
                                          BranchNode_items_start_timeZone,
                                          BranchNode_items_end_date,
                                          BranchNode_items_end_dateTime,
                                          BranchNode_items_end_timeZone,
                                          BranchNode_items_endTimeUnspecified,
                                          BranchNode_items_recurrence,
                                          BranchNode_items_recurringEventId,
                                          BranchNode_items_originalStartTime_date,
                                          BranchNode_items_originalStartTime_dateTime,
                                          BranchNode_items_originalStartTime_timeZone,
                                          BranchNode_items_transparency,
                                          BranchNode_items_visibility,
                                          BranchNode_items_iCalUID,
                                          BranchNode_items_sequence,
                                          BranchNode_items_attendees,
                                          BranchNode_items_attendeesOmitted,
                                          BranchNode_items_extendedProperties_private,
                                          BranchNode_items_extendedProperties_shared,
                                          BranchNode_items_hangoutLink,
                                          BranchNode_items_conferenceData_createRequest_requestId,
                                          BranchNode_items_conferenceData_createRequest_conferenceSolutionKey_type,
                                          BranchNode_items_conferenceData_createRequest_status_statusCode,
                                          BranchNode_items_conferenceData_entryPoints,
                                          BranchNode_items_conferenceData_conferenceSolution_key_type,
                                          BranchNode_items_conferenceData_conferenceSolution_name,
                                          BranchNode_items_conferenceData_conferenceSolution_iconUri,
                                          BranchNode_items_conferenceData_conferenceId,
                                          BranchNode_items_conferenceData_signature,
                                          BranchNode_items_conferenceData_notes,
                                          BranchNode_items_gadget_type,
                                          BranchNode_items_gadget_title,
                                          BranchNode_items_gadget_link,
                                          BranchNode_items_gadget_iconLink,
                                          BranchNode_items_gadget_width,
                                          BranchNode_items_gadget_height,
                                          BranchNode_items_gadget_display,
                                          BranchNode_items_gadget_preferences,
                                          BranchNode_items_anyoneCanAddSelf,
                                          BranchNode_items_guestsCanInviteOthers,
                                          BranchNode_items_guestsCanModify,
                                          BranchNode_items_guestsCanSeeOtherGuests,
                                          BranchNode_items_privateCopy,
                                          BranchNode_items_locked,
                                          BranchNode_reminders_useDefault,
                                          BranchNode_reminders_overrides,
                                          BranchNode_source_url,
                                          BranchNode_source_title,
                                          BranchNode_attachments
                                          ]

                    EveryBranchNodeItem = [removeBlank(str(ele)) for ele in EveryBranchNodeItem]

                    Add_ele_to_BranchNodeItem = [RootNode_kind,
                                                RootNode_etag,
                                                RootNode_summary,
                                                RootNode_description,
                                                RootNode_updated,
                                                RootNode_timeZone,
                                                RootNode_accessRole,
                                                RootNode_defaultReminders,
                                                RootNode_nextPageToken,
                                                RootNode_nextSyncToken,
                                                *EveryBranchNodeItem]

                    List_Root_And_Branch.append(Add_ele_to_BranchNodeItem)

        Need_To_Save_List = [] #保存昨天(东京时间)00:00:00.000~23:59:59.999的数据
        for item in List_Root_And_Branch:
            if (item[15] >=TwoDaysAgoDate + UTCStartTime and item[15] <= Yesterday + UTCEndTime)  or (item[16] >=TwoDaysAgoDate + UTCStartTime and item[16] <= Yesterday + UTCEndTime): #如果创建时间大于昨天,就删除
                Need_To_Save_List.append(item)

        return Need_To_Save_List


def save_txt_to_disk(paraCalendarPath,paraYesterday,paraSaveList):
   '''
   :param paraCalendarPath: 日历目录
   :param paraYesterday: 昨天日期(年-月-日)
   :param paraSaveList: 要保存的文本的List
   :return: 
   '''
   try:
       # if not os.path.exists(paraCalendarPath + paraYesterday):
       #     os.makedirs(paraCalendarPath + paraYesterday)
       if paraSaveList:
           if not os.path.exists(paraCalendarPath):
               os.makedirs(paraCalendarPath)
           with open(paraCalendarPath + paraYesterday + '.txt', "w", encoding="utf-8") as fo:
               fo.write('\n'.join([' '.join(i) for i in paraSaveList]))
   except Exception as ex:
       logger.error("Call method save_txt_to_disk() error!")
       raise ex




if __name__ == '__main__':
    logger = write_log()  # 获取日志对象
    time_start = datetime.datetime.now()
    start = time.time()
    logger.info("Program start,now time is:" + str(time_start))
    SCOPES = 'https://www.googleapis.com/auth/calendar'
    TwoDaysAgoDate = datetime.datetime.now() - datetime.timedelta(days=2)  # 前天日期
    TwoDaysAgoDate = TwoDaysAgoDate.strftime('%Y-%m-%d')
    Yesterday = datetime.datetime.now() - datetime.timedelta(days=1) #昨天日期
    Yesterday = Yesterday.strftime('%Y-%m-%d')
    CalendarPath = r'./Calendar/'#路径
    gmailrList = getcalendarIdList()
    Root_And_Branch_To_Txt = generateEveryDayCalendarData(gmailrList)
    save_txt_to_disk(CalendarPath, Yesterday, Root_And_Branch_To_Txt)
    time_end = datetime.datetime.now()
    end = time.time()
    logger.info("Program end,now time is:" + str(time_end))
    logger.info("Program run : %f seconds" % (end - start))

