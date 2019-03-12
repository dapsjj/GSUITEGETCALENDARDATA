#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import print_function
import time
import logging
import os
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import socket
import re
import csv
import configparser


CalendarIdListFile = None
TokenFile = None
CalendarAdministratorFile = None
SCOPES = None
CalendarPath = None
EventTimeMin = None
EventTimeMax = None
HowManyDays = None
EventTextName = None
CreateTextName = None
allEvents = None
summary = None
TokyoStartTime = None
UTCStartTime = None
UTCEndTime = None

socket.setdefaulttimeout(300)

def write_log():
    '''
    写log
    :return: 返回logger对象
    '''
    # 获取logger实例，如果参数为空则返回root logger
    logger = logging.getLogger()
    now_date = datetime.datetime.now().strftime('%Y%m%d')
    log_file = now_date + ".log"  # 文件日志
    if not os.path.exists("log"):  # python文件同级别创建log文件夹
        os.makedirs("log")
    # 指定logger输出格式
    formatter = logging.Formatter('%(asctime)s %(levelname)s line:%(lineno)s %(message)s')
    file_handler = logging.FileHandler("log" + os.sep + log_file, mode='a', encoding='utf-8')
    file_handler.setFormatter(formatter)  # 可以通过setFormatter指定输出格式
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
    try:
        MyString = re.sub('[\s+]', '', MyString)
        return MyString
    except Exception as ex:
        logger.error("Call method removeBlank() error!")
        raise ex


def getcalendarIdList():
    '''
    读人员邮箱地址
    '''
    try:
        txt_config_list = []
        fileName = CalendarIdListFile
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


def generateEveryDayCalendarData(GmailList, paraTwoDaysAgoDate, paraYesterday):
    '''
    :param GmailList: API中calendarId对应的List
    :param paraTwoDaysAgoDate: 调用api的updatedMin创建或者更新的开始日期)
    :param paraYesterday: 事件的创建或者更新的结束日期
    :return: 要保存成文本的List

    从根节点到分支节点(叶子节点)的全部元素的说明
    共79个元素:
    1、RootNode_kind:集合的类型("calendar#events")
    2、RootNode_etag:ETag的集合
    3、RootNode_summary:日历的标题
    4、RootNode_description:日历说明
    5、RootNode_updated:日历的上次修改时间
    6、RootNode_timeZone:日历的时区
    7、RootNode_accessRole:用户对此日历的访问角色
    8、RootNode_defaultReminders:已验证用户的日历上的默认提醒
    9、RootNode_nextPageToken:令牌用于访问此结果的下一页。如果没有进一步的结果，则省略
    10、RootNode_nextSyncToken:在稍后的时间点使用的令牌仅检索自返回此结果以来已更改的条目
    11、BranchNode_items_kind:资源的类型("calendar#event")
    12、BranchNode_items_etag:etag的资源
    13、BranchNode_items_id: 事件的不透明标识符
    14、BranchNode_items_status: 事件的状态
    15、BranchNode_items_htmlLink:Google Calendar Web UI中此事件的绝对链接 
    16、BranchNode_items_created:事件的创建时间
    17、BranchNode_items_updated:事件的上次修改时间
    18、BranchNode_items_summary:活动的标题
    19、BranchNode_items_description: 事件描述
    20、BranchNode_items_location: 事件的地理位置为自由格式文本
    21、BranchNode_items_colorId: 事件的颜色
    22、BranchNode_items_creator_id:创建者的个人资料ID
    23、BranchNode_items_creator_email:创建者的电子邮件地址
    24、BranchNode_items_creator_displayName:创建者的姓名
    25、BranchNode_items_creator_self:创建者是否对应于此事件副本出现的日历
    26、BranchNode_items_organizer_id:组织者的个人资料ID
    27、BranchNode_items_organizer_email:组织者的电子邮件地址
    28、BranchNode_items_organizer_displayName:组织者的名称
    29、BranchNode_items_organizer_self:组织者是否对应于此事件副本出现的日历
    30、BranchNode_items_start_date:事件开始日期,格式为"yyyy-mm-dd"
    31、BranchNode_items_start_dateTime:事件开始时间
    32、BranchNode_items_start_timeZone:事件开始的时区
    33、BranchNode_items_end_date:事件结束日期,格式为"yyyy-mm-dd"
    34、BranchNode_items_end_dateTime:事件结束时间
    35、BranchNode_items_end_timeZone:事件结束时区
    36、BranchNode_items_endTimeUnspecified:结束时间是否实际未指定
    37、BranchNode_items_recurrence:RFC5545中指定的循环事件的RRULE，EXRULE，RDATE和EXDATE行列表
    38、BranchNode_items_recurringEventId:对于周期性事件的实例
    39、BranchNode_items_originalStartTime_date:原始开始日期
    40、BranchNode_items_originalStartTime_dateTime:原始开始时间
    41、BranchNode_items_originalStartTime_timeZone :原始开始时间时区
    42、BranchNode_items_transparency:事件是否会阻止日历上的时间
    43、BranchNode_items_visibility:活动的可见性
    44、BranchNode_items_iCalUID:RFC5545中定义的事件唯一标识符
    45、BranchNode_items_sequence:根据iCalendar的序列号
    46、BranchNode_items_attendees:活动的与会者
    47、BranchNode_items_attendeesOmitted: 参与者是否可能已从事件的表示中省略
    48、BranchNode_items_extendedProperties_private:属于此日历上显示的事件副本的属性
    49、BranchNode_items_extendedProperties_shared:在其他与会者日历上的事件副本之间共享的属性
    50、BranchNode_items_hangoutLink:与此活动相关联的Google+视频群聊的绝对链接
    51、BranchNode_items_conferenceData_createRequest_requestId:客户端为此请求生成的唯一ID
    52、BranchNode_items_conferenceData_createRequest_conferenceSolutionKey_type:会议解决方案类型
    53、BranchNode_items_conferenceData_createRequest_status_statusCode:会议的当前状态创建请求
    54、BranchNode_items_conferenceData_entryPoints:有关各个会议入口点的信息
    55、BranchNode_items_conferenceData_conferenceSolution_key_type:会议解决方案类型
    56、BranchNode_items_conferenceData_conferenceSolution_name:此解决方案的用户可见名称
    57、BranchNode_items_conferenceData_conferenceSolution_iconUri:此解决方案的用户可见图标
    58、BranchNode_items_conferenceData_conferenceId:会议的ID
    59、BranchNode_items_conferenceData_signature:会议数据的签名
    60、BranchNode_items_conferenceData_notes:要向用户显示的其他说明
    61、BranchNode_items_gadget_type:小工具的类型
    62、BranchNode_items_gadget_title:小工具的标题
    63、BranchNode_items_gadget_link:小工具的网址
    64、BranchNode_items_gadget_iconLink:小工具的图标网址
    65、BranchNode_items_gadget_width:小工具的宽度
    66、BranchNode_items_gadget_height:小工具的高度
    67、BranchNode_items_gadget_display:小工具的显示模式
    68、BranchNode_items_gadget_preferences:首选项名称和相应的值
    69、BranchNode_items_anyoneCanAddSelf:是否有人可以邀请自己参加此活动
    70、BranchNode_items_guestsCanInviteOthers:组织者以外的与会者是否可以邀请其他人参加活动
    71、BranchNode_items_guestsCanModify:组织者以外的与会者是否可以修改活动
    72、BranchNode_items_guestsCanSeeOtherGuests:组织者以外的与会者是否可以看到活动的与会者是谁
    73、BranchNode_items_privateCopy:这是否是私有事件副本
    74、BranchNode_items_locked:这是否是锁定事件副本
    75、BranchNode_reminders_useDefault:是否将日历的默认提醒应用于该事件
    76、BranchNode_reminders_overrides:如果事件未使用默认提醒，则会列出特定于事件的提醒
    77、BranchNode_source_url:指向资源的源的URL
    78、BranchNode_source_title:来源名称
    79、BranchNode_attachments:该事件的文件附件
    80、calendarId:自己添加的calendarId为了以后匹配人名用

    '''
    try:
        if GmailList:
            store = file.Storage(TokenFile)  # 如果换了证书的json文件，需要是删除token.json
            creds = store.get()
            if not creds or creds.invalid:
                flow = client.flow_from_clientsecrets(CalendarAdministratorFile, SCOPES)
                creds = tools.run_flow(flow, store)
            service = build('calendar', 'v3', http=creds.authorize(Http()))
            UTCStartTime = 'T15:00:00.000Z'  # UTC时间
            UTCEndTime = 'T14:59:59.999Z'  # UTC时间

            List_Root_And_Branch = []  # Root节点和Branch节点拼接到一起
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
                                                      singleEvents=True,
                                                      updatedMin=paraTwoDaysAgoDate + UTCStartTime
                                                      # 相当于日本时间昨天00:00:00.000
                                                      ).execute()

                RootNode_kind = removeBlank(str(events_result.get('kind', '_')))
                RootNode_etag = removeBlank(str(events_result.get('etag', '_')))
                RootNode_summary = removeBlank(str(events_result.get('summary', '_')))
                RootNode_description = removeBlank(str(events_result.get('description', '_')))
                RootNode_updated = removeBlank(str(events_result.get('updated', '_')))
                RootNode_timeZone = removeBlank(str(events_result.get('timeZone', '_')))
                RootNode_accessRole = removeBlank(str(events_result.get('accessRole', '_')))
                RootNode_defaultReminders = removeBlank(str(events_result.get('defaultReminders', '[]')))  # 数组
                RootNode_nextPageToken = removeBlank(str(events_result.get('nextPageToken', '_')))
                RootNode_nextSyncToken = removeBlank(str(events_result.get('nextSyncToken', '_')))
                RootNode_items = events_result.get('items')  # Root节点下的items节点,是个数组

                if RootNode_items:  # 判断items是否为空
                    for item in RootNode_items:
                        EveryBranchNodeItem = []  # 每个items中的节点
                        Add_ele_to_BranchNodeItem = []  # 给EveryBranchNodeItem添加前面的元素用的List
                        BranchNode_items_kind = item.get('kind', '_')
                        BranchNode_items_etag = item.get('etag', '_')
                        BranchNode_items_id = item.get('id', '_')
                        BranchNode_items_status = item.get('status', '_')
                        BranchNode_items_htmlLink = item.get('htmlLink', '_')
                        BranchNode_items_created = item.get('created', '_')
                        BranchNode_items_updated = item.get('updated', '_')
                        BranchNode_items_summary = item.get('summary', '_')
                        BranchNode_items_description = item.get('description', '_')
                        BranchNode_items_location = item.get('location', '_')
                        BranchNode_items_colorId = item.get('colorId', '_')
                        # BranchNode_items_creator = item.get('creator','_')
                        BranchNode_items_creator_id = item.get('creator', {}).get('id', '_')
                        BranchNode_items_creator_email = item.get('creator', {}).get('email', '_')
                        BranchNode_items_creator_displayName = item.get('creator', {}).get('displayName', '_')
                        BranchNode_items_creator_self = item.get('creator', {}).get('self', '_')
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
                        BranchNode_items_endTimeUnspecified = item.get('endTimeUnspecified', '_')
                        BranchNode_items_recurrence = item.get('recurrence', '[]')  # 数组
                        BranchNode_items_recurringEventId = item.get('recurringEventId', '_')
                        BranchNode_items_originalStartTime_date = item.get('originalStartTime', {}).get('date', '_')
                        BranchNode_items_originalStartTime_dateTime = item.get('originalStartTime', {}).get('dateTime',
                                                                                                            '_')
                        BranchNode_items_originalStartTime_timeZone = item.get('originalStartTime', {}).get('timeZone',
                                                                                                            '_')
                        BranchNode_items_transparency = item.get('transparency', '_')
                        BranchNode_items_visibility = item.get('visibility', '_')
                        BranchNode_items_iCalUID = item.get('iCalUID', '_')
                        BranchNode_items_sequence = item.get('sequence', '_')
                        BranchNode_items_attendees = item.get('attendees', '[]')
                        BranchNode_items_attendeesOmitted = item.get('attendeesOmitted', '_')
                        # BranchNode_items_extendedProperties = item.get('extendedProperties','_')
                        BranchNode_items_extendedProperties_private = item.get('extendedProperties', {}).get('private',
                                                                                                             '_')
                        BranchNode_items_extendedProperties_shared = item.get('extendedProperties', {}).get('shared',
                                                                                                            '_')
                        BranchNode_items_hangoutLink = item.get('hangoutLink', '_')
                        # BranchNode_items_conferenceData = item.get('conferenceData','_')
                        BranchNode_items_conferenceData_createRequest_requestId = item.get('conferenceData', {}).get(
                            'createRequest', {}).get('requestId', '_')
                        BranchNode_items_conferenceData_createRequest_conferenceSolutionKey_type = item.get(
                            'conferenceData', {}).get('createRequest', {}).get('conferenceSolutionKey', {}).get('type',
                                                                                                                '_')
                        BranchNode_items_conferenceData_createRequest_status_statusCode = item.get('conferenceData',
                                                                                                   {}).get(
                            'createRequest', {}).get('status', {}).get('statusCode', '_')
                        BranchNode_items_conferenceData_entryPoints = item.get('conferenceData', {}).get('entryPoints',
                                                                                                         '[]')  # 数组
                        # BranchNode_items_conferenceSolution = item.get('conferenceData',{}).get('conferenceSolution','_')
                        BranchNode_items_conferenceData_conferenceSolution_key_type = item.get('conferenceData',
                                                                                               {}).get(
                            'conferenceSolution', {}).get('key', {}).get('type', '_')
                        BranchNode_items_conferenceData_conferenceSolution_name = item.get('conferenceData', {}).get(
                            'conferenceSolution', {}).get('name', '_')
                        BranchNode_items_conferenceData_conferenceSolution_iconUri = item.get('conferenceData', {}).get(
                            'conferenceSolution', {}).get('iconUri', '_')
                        BranchNode_items_conferenceData_conferenceId = item.get('conferenceData', {}).get(
                            'conferenceId', '_')
                        BranchNode_items_conferenceData_signature = item.get('conferenceData', {}).get('signature', '_')
                        BranchNode_items_conferenceData_notes = item.get('conferenceData', {}).get('notes', '_')
                        # BranchNode_items_gadget = item.get('gadget','_')
                        BranchNode_items_gadget_type = item.get('gadget', {}).get('type', '_')
                        BranchNode_items_gadget_title = item.get('gadget', {}).get('title', '_')
                        BranchNode_items_gadget_link = item.get('gadget', {}).get('link', '_')
                        BranchNode_items_gadget_iconLink = item.get('gadget', {}).get('iconLink', '_')
                        BranchNode_items_gadget_width = item.get('gadget', {}).get('width', '_')
                        BranchNode_items_gadget_height = item.get('gadget', {}).get('height', '_')
                        BranchNode_items_gadget_display = item.get('gadget', {}).get('display', '_')
                        BranchNode_items_gadget_preferences = item.get('gadget', {}).get('preferences', '_')
                        BranchNode_items_anyoneCanAddSelf = item.get('anyoneCanAddSelf', '_')
                        BranchNode_items_guestsCanInviteOthers = item.get('guestsCanInviteOthers', '_')
                        BranchNode_items_guestsCanModify = item.get('guestsCanModify', '_')
                        BranchNode_items_guestsCanSeeOtherGuests = item.get('guestsCanSeeOtherGuests', '_')
                        BranchNode_items_privateCopy = item.get('privateCopy', '_')
                        BranchNode_items_locked = item.get('locked', '_')
                        # BranchNode_reminders = item.get('reminders','_')
                        BranchNode_reminders_useDefault = item.get('reminders', {}).get('useDefault', '_')
                        BranchNode_reminders_overrides = item.get('reminders', {}).get('overrides', '[]')  # 数组
                        # BranchNode_source = item.get('source','_')
                        BranchNode_source_url = item.get('source', {}).get('url', '_')
                        BranchNode_source_title = item.get('source', {}).get('title', '_')
                        BranchNode_attachments = item.get('attachments', '[]')  # 数组

                        # 添加到分支节点的List中
                        EveryBranchNodeItem = [BranchNode_items_kind,
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
                                               BranchNode_attachments,
                                               calendarId  # 自己添加的calendarId为了以后匹配人名用
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

            Need_To_Save_List = []  # 保存昨天(东京时间)00:00:00.000~23:59:59.999的数据
            for item in List_Root_And_Branch:
                if (item[15] >= paraTwoDaysAgoDate + UTCStartTime and item[15] <= paraYesterday + UTCEndTime) \
                        or (item[16] >= paraTwoDaysAgoDate + UTCStartTime and item[16] <= paraYesterday + UTCEndTime):  # 如果创建时间大于昨天,就删除
                    Need_To_Save_List.append(item)

            return Need_To_Save_List

    except Exception as ex:
        logger.error("Call method generateEveryDayCalendarData() error!")
        raise ex


def save_txt_to_disk(paraCalendarPath, paraYesterday, paraSaveList):
    '''
    :param paraCalendarPath: 日历目录
    :param paraYesterday: 昨天日期(年-月-日)
    :param paraSaveList: 要保存的文本的List
    :return: 
    '''
    try:
        if paraSaveList:
            if not os.path.exists(paraCalendarPath):
                os.makedirs(paraCalendarPath)
            with open(paraCalendarPath + paraYesterday + '.txt', "w", encoding="utf-8") as fo:
                fo.write('\n'.join([' '.join(i) for i in paraSaveList]))
    except Exception as ex:
        logger.error("Call method save_txt_to_disk() error!")
        raise ex


def save_data_to_csv(paraCalendarPath, paraYesterday, paraSaveList):
    '''
    :param paraCalendarPath: 日历目录
    :param paraYesterday: 昨天日期(年-月-日)
    :param paraSaveList: 要保存的文本的List
    :return: 
    '''
    title = [['RootNode_kind', 'RootNode_etag', 'RootNode_summary', 'RootNode_description', 'RootNode_updated',
              'RootNode_timeZone', 'RootNode_accessRole', 'RootNode_defaultReminders', 'RootNode_nextPageToken',
              'RootNode_nextSyncToken', 'BranchNode_items_kind', 'BranchNode_items_etag', 'BranchNode_items_id',
              'BranchNode_items_status', 'BranchNode_items_htmlLink', 'BranchNode_items_created',
              'BranchNode_items_updated', 'BranchNode_items_summary', 'BranchNode_items_description',
              'BranchNode_items_location', 'BranchNode_items_colorId', 'BranchNode_items_creator_id',
              'BranchNode_items_creator_email', 'BranchNode_items_creator_displayName', 'BranchNode_items_creator_self',
              'BranchNode_items_organizer_id', 'BranchNode_items_organizer_email',
              'BranchNode_items_organizer_displayName', 'BranchNode_items_organizer_self',
              'BranchNode_items_start_date', 'BranchNode_items_start_dateTime', 'BranchNode_items_start_timeZone',
              'BranchNode_items_end_date', 'BranchNode_items_end_dateTime', 'BranchNode_items_end_timeZone',
              'BranchNode_items_endTimeUnspecified', 'BranchNode_items_recurrence', 'BranchNode_items_recurringEventId',
              'BranchNode_items_originalStartTime_date', 'BranchNode_items_originalStartTime_dateTime',
              'BranchNode_items_originalStartTime_timeZone', 'BranchNode_items_transparency',
              'BranchNode_items_visibility', 'BranchNode_items_iCalUID', 'BranchNode_items_sequence',
              'BranchNode_items_attendees', 'BranchNode_items_attendeesOmitted',
              'BranchNode_items_extendedProperties_private', 'BranchNode_items_extendedProperties_shared',
              'BranchNode_items_hangoutLink', 'BranchNode_items_conferenceData_createRequest_requestId',
              'BranchNode_items_conferenceData_createRequest_conferenceSolutionKey_type',
              'BranchNode_items_conferenceData_createRequest_status_statusCode',
              'BranchNode_items_conferenceData_entryPoints',
              'BranchNode_items_conferenceData_conferenceSolution_key_type',
              'BranchNode_items_conferenceData_conferenceSolution_name',
              'BranchNode_items_conferenceData_conferenceSolution_iconUri',
              'BranchNode_items_conferenceData_conferenceId', 'BranchNode_items_conferenceData_signature',
              'BranchNode_items_conferenceData_notes', 'BranchNode_items_gadget_type', 'BranchNode_items_gadget_title',
              'BranchNode_items_gadget_link', 'BranchNode_items_gadget_iconLink', 'BranchNode_items_gadget_width',
              'BranchNode_items_gadget_height', 'BranchNode_items_gadget_display',
              'BranchNode_items_gadget_preferences', 'BranchNode_items_anyoneCanAddSelf',
              'BranchNode_items_guestsCanInviteOthers', 'BranchNode_items_guestsCanModify',
              'BranchNode_items_guestsCanSeeOtherGuests', 'BranchNode_items_privateCopy', 'BranchNode_items_locked',
              'BranchNode_reminders_useDefault', 'BranchNode_reminders_overrides', 'BranchNode_source_url',
              'BranchNode_source_title', 'BranchNode_attachments', 'gmail']]
    try:
        if paraSaveList:
            if not os.path.exists(paraCalendarPath):
                os.makedirs(paraCalendarPath)
            with open(paraCalendarPath + paraYesterday + '.csv', "w", newline='', encoding="utf-8") as fo:
                writer = csv.writer(fo)
                writer.writerows(title)
                writer.writerows(paraSaveList)
    except Exception as ex:
        logger.error("Call method save_data_to_csv() error!")
        raise ex


def getEveryDay(begin_date, end_date):
    try:
        date_list = []
        begin_date = datetime.datetime.strptime(begin_date, "%Y-%m-%d")
        end_date = datetime.datetime.strptime(end_date, "%Y-%m-%d")
        while begin_date <= end_date:
            date_str = begin_date.strftime("%Y-%m-%d")
            date_list.append(date_str)
            begin_date += datetime.timedelta(days=1)
        return date_list
    except Exception as ex:
        logger.error("Call method getEveryDay() error!")
        raise ex


def generateSummaryDate(paraAllDateList):
    '''
    :param paraAllDateList: 要处理的日期列表
    :return: 
    '''
    try:
        if paraAllDateList:
            allDataList = []  # 存放全部数据的List

            for everyDate in paraAllDateList:
                if os.path.exists(CalendarPath + everyDate + '.txt'):
                    currentDataList = []  # 存放当前遍历的数据List
                    with open(CalendarPath + everyDate + '.txt', 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                        for line in lines:
                            if not line:  # 如果line是空
                                continue
                            else:
                                row_list = line.split(" ")
                                currentDataList.append(row_list)
                    for row in currentDataList:
                        if not allDataList:  # 如果存放全部数据的List是空
                            if row[13] != 'cancelled':
                                allDataList.append(row)
                        else:  # 如果存放全部数据的List不是空
                            iCalUIDList = [x[43] for x in allDataList]
                            if row[13] != 'cancelled':  # 不是删除的事件
                                if row[43] not in iCalUIDList:  # 没有这个iCalUID
                                    allDataList.append(row)
                                else:  # 有这个iCalUID
                                    findIndex = iCalUIDList.index(row[43])
                                    # 删除原来的iCalUID对应的数据,插入该iCalUID对应的新数据(status='tentative' or status='confirmed')
                                    del allDataList[findIndex]
                                    allDataList.append(row)
                            else:  # 是删除的事件
                                if row[43] in iCalUIDList:
                                    delfindIndex = iCalUIDList.index(row[43])
                                    # 删除原来的iCalUID对应的数据
                                    del allDataList[delfindIndex]

                    # 处理事件的日期
                    for item in allDataList:
                        if item[29] == '_':  # 不是全天事件,则事件的开始时间有'年月日'和'时分秒'
                            eventStartDate = item[30][0:10]  # 取开始时间的年月日
                            eventStartTime = item[30][11:19]  # 取开始时间的时分秒
                        else:  # 是全天事件,则事件的开始时间只有年月日
                            eventStartDate = item[29][0:10]  # 取开始时间的年月日
                            eventStartTime = ''

                        if item[32] == '_':  # 不是全天事件,则事件的结束时间有'年月日'和'时分秒'
                            eventEndDate = item[33][0:10]  # 取结束时间的年月日
                            eventEndtTime = item[33][11:19]  # 取结束时间的时分秒
                        else:  # 是全天事件,则事件的结束时间只有年月日
                            eventEndDate = item[32][0:10]  # 取开始时间的年月日
                            eventEndtTime = ''

    except Exception as ex:
        logger.error("Call method generateAllDataByStartDateAndEndDate() error!")
        raise ex


def SaveEveryDayCalendarDataUsetimeMin_timeMax(paraEventDateList, paraGmailList):
    try:
        len_EventDateList = len(paraEventDateList)
        if len_EventDateList > 1 and paraGmailList:
            if not os.path.exists(CalendarPath):
                os.makedirs(CalendarPath)

            for i in range(len_EventDateList - 1):
                List_Root_And_Branch = []  # Root节点和Branch节点拼接到一起
                EventStartDate = paraEventDateList[i]
                EventEndDate = paraEventDateList[i + 1]
                Need_To_Save_List = []  # 保存某一天的文本数据
                for calendarId in paraGmailList:
                    store = file.Storage(TokenFile)  # 如果换了证书的json文件，需要是删除token.json
                    creds = store.get()
                    if not creds or creds.invalid:
                        flow = client.flow_from_clientsecrets(CalendarAdministratorFile, SCOPES)
                        creds = tools.run_flow(flow, store)
                    service = build('calendar', 'v3', http=creds.authorize(Http()))
                    events_result = service.events().list(calendarId=calendarId,
                                                          orderBy='updated',  # orderBy允许的值['startTime', 'updated']
                                                          showHiddenInvitations=True,
                                                          showDeleted=True,
                                                          singleEvents=True, #重复事件以单独的事件的形式返回
                                                          timeMin=EventStartDate + TokyoStartTime,
                                                          timeMax=EventEndDate + TokyoStartTime
                                                          ).execute()

                    RootNode_kind = removeBlank(str(events_result.get('kind', '_')))
                    RootNode_etag = removeBlank(str(events_result.get('etag', '_')))
                    RootNode_summary = removeBlank(str(events_result.get('summary', '_')))
                    RootNode_description = removeBlank(str(events_result.get('description', '_')))
                    RootNode_updated = removeBlank(str(events_result.get('updated', '_')))
                    RootNode_timeZone = removeBlank(str(events_result.get('timeZone', '_')))
                    RootNode_accessRole = removeBlank(str(events_result.get('accessRole', '_')))
                    RootNode_defaultReminders = removeBlank(str(events_result.get('defaultReminders', '[]')))  # 数组
                    RootNode_nextPageToken = removeBlank(str(events_result.get('nextPageToken', '_')))
                    RootNode_nextSyncToken = removeBlank(str(events_result.get('nextSyncToken', '_')))
                    RootNode_items = events_result.get('items')  # Root节点下的items节点,是个数组

                    if RootNode_items:  # 判断items是否为空
                        for item in RootNode_items:
                            EveryBranchNodeItem = []  # 每个items中的节点
                            Add_ele_to_BranchNodeItem = []  # 给EveryBranchNodeItem添加前面的元素用的List
                            BranchNode_items_kind = item.get('kind', '_')
                            BranchNode_items_etag = item.get('etag', '_')
                            BranchNode_items_id = item.get('id', '_')
                            BranchNode_items_status = item.get('status', '_')
                            BranchNode_items_htmlLink = item.get('htmlLink', '_')
                            BranchNode_items_created = item.get('created', '_')
                            BranchNode_items_updated = item.get('updated', '_')
                            BranchNode_items_summary = item.get('summary', '_')
                            BranchNode_items_description = item.get('description', '_')
                            BranchNode_items_location = item.get('location', '_')
                            BranchNode_items_colorId = item.get('colorId', '_')
                            # BranchNode_items_creator = item.get('creator','_')
                            BranchNode_items_creator_id = item.get('creator', {}).get('id', '_')
                            BranchNode_items_creator_email = item.get('creator', {}).get('email', '_')
                            BranchNode_items_creator_displayName = item.get('creator', {}).get('displayName', '_')
                            BranchNode_items_creator_self = item.get('creator', {}).get('self', '_')
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
                            BranchNode_items_endTimeUnspecified = item.get('endTimeUnspecified', '_')
                            BranchNode_items_recurrence = item.get('recurrence', '[]')  # 数组
                            BranchNode_items_recurringEventId = item.get('recurringEventId', '_')
                            BranchNode_items_originalStartTime_date = item.get('originalStartTime', {}).get('date', '_')
                            BranchNode_items_originalStartTime_dateTime = item.get('originalStartTime', {}).get(
                                'dateTime', '_')
                            BranchNode_items_originalStartTime_timeZone = item.get('originalStartTime', {}).get(
                                'timeZone', '_')
                            BranchNode_items_transparency = item.get('transparency', '_')
                            BranchNode_items_visibility = item.get('visibility', '_')
                            BranchNode_items_iCalUID = item.get('iCalUID', '_')
                            BranchNode_items_sequence = item.get('sequence', '_')
                            BranchNode_items_attendees = item.get('attendees', '[]')
                            BranchNode_items_attendeesOmitted = item.get('attendeesOmitted', '_')
                            # BranchNode_items_extendedProperties = item.get('extendedProperties','_')
                            BranchNode_items_extendedProperties_private = item.get('extendedProperties', {}).get(
                                'private', '_')
                            BranchNode_items_extendedProperties_shared = item.get('extendedProperties', {}).get(
                                'shared', '_')
                            BranchNode_items_hangoutLink = item.get('hangoutLink', '_')
                            # BranchNode_items_conferenceData = item.get('conferenceData','_')
                            BranchNode_items_conferenceData_createRequest_requestId = item.get('conferenceData',
                                                                                               {}).get('createRequest',
                                                                                                       {}).get(
                                'requestId', '_')
                            BranchNode_items_conferenceData_createRequest_conferenceSolutionKey_type = item.get(
                                'conferenceData', {}).get('createRequest', {}).get('conferenceSolutionKey', {}).get(
                                'type', '_')
                            BranchNode_items_conferenceData_createRequest_status_statusCode = item.get('conferenceData',
                                                                                                       {}).get(
                                'createRequest', {}).get('status', {}).get('statusCode', '_')
                            BranchNode_items_conferenceData_entryPoints = item.get('conferenceData', {}).get(
                                'entryPoints', '[]')  # 数组
                            # BranchNode_items_conferenceSolution = item.get('conferenceData',{}).get('conferenceSolution','_')
                            BranchNode_items_conferenceData_conferenceSolution_key_type = item.get('conferenceData',
                                                                                                   {}).get(
                                'conferenceSolution', {}).get('key', {}).get('type', '_')
                            BranchNode_items_conferenceData_conferenceSolution_name = item.get('conferenceData',
                                                                                               {}).get(
                                'conferenceSolution', {}).get('name', '_')
                            BranchNode_items_conferenceData_conferenceSolution_iconUri = item.get('conferenceData',
                                                                                                  {}).get(
                                'conferenceSolution', {}).get('iconUri', '_')
                            BranchNode_items_conferenceData_conferenceId = item.get('conferenceData', {}).get(
                                'conferenceId', '_')
                            BranchNode_items_conferenceData_signature = item.get('conferenceData', {}).get('signature',
                                                                                                           '_')
                            BranchNode_items_conferenceData_notes = item.get('conferenceData', {}).get('notes', '_')
                            # BranchNode_items_gadget = item.get('gadget','_')
                            BranchNode_items_gadget_type = item.get('gadget', {}).get('type', '_')
                            BranchNode_items_gadget_title = item.get('gadget', {}).get('title', '_')
                            BranchNode_items_gadget_link = item.get('gadget', {}).get('link', '_')
                            BranchNode_items_gadget_iconLink = item.get('gadget', {}).get('iconLink', '_')
                            BranchNode_items_gadget_width = item.get('gadget', {}).get('width', '_')
                            BranchNode_items_gadget_height = item.get('gadget', {}).get('height', '_')
                            BranchNode_items_gadget_display = item.get('gadget', {}).get('display', '_')
                            BranchNode_items_gadget_preferences = item.get('gadget', {}).get('preferences', '_')
                            BranchNode_items_anyoneCanAddSelf = item.get('anyoneCanAddSelf', '_')
                            BranchNode_items_guestsCanInviteOthers = item.get('guestsCanInviteOthers', '_')
                            BranchNode_items_guestsCanModify = item.get('guestsCanModify', '_')
                            BranchNode_items_guestsCanSeeOtherGuests = item.get('guestsCanSeeOtherGuests', '_')
                            BranchNode_items_privateCopy = item.get('privateCopy', '_')
                            BranchNode_items_locked = item.get('locked', '_')
                            # BranchNode_reminders = item.get('reminders','_')
                            BranchNode_reminders_useDefault = item.get('reminders', {}).get('useDefault', '_')
                            BranchNode_reminders_overrides = item.get('reminders', {}).get('overrides', '[]')  # 数组
                            # BranchNode_source = item.get('source','_')
                            BranchNode_source_url = item.get('source', {}).get('url', '_')
                            BranchNode_source_title = item.get('source', {}).get('title', '_')
                            BranchNode_attachments = item.get('attachments', '[]')  # 数组

                            # 添加到分支节点的List中
                            EveryBranchNodeItem = [BranchNode_items_kind,
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
                                                   BranchNode_attachments,
                                                   calendarId  # 自己添加的calendarId为了以后匹配人名用
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

                for item in List_Root_And_Branch:
                    if item[29] == '_':  # 不是全天事件,则事件的开始时间有'年月日'和'时分秒'
                        temp_eventStartDate = item[30][0:10]  # 取开始时间的年月日
                    else:  # 是全天事件,则事件的开始时间只有年月日
                        temp_eventStartDate = item[29][0:10]  # 取开始时间的年月日
                    if item[32] == '_':  # 不是全天事件,则事件的结束时间有'年月日'和'时分秒'
                        temp_eventEndDate = item[33][0:10]  # 取结束时间的年月日
                    else:  # 是全天事件,则事件的结束时间只有年月日
                        temp_eventEndDate = item[32][0:10]  # 取开始时间的年月日
                    if (temp_eventStartDate >= EventStartDate and temp_eventStartDate <= EventEndDate) or (
                            temp_eventEndDate >= EventStartDate and temp_eventEndDate < EventEndDate):  # 如果事件的开始和结束时间超过指定的开始和结束时间,就删除.连续事件如果用temp_eventEndDate <= EventEndDate，则在结束的时候会连续出现2次，因此改成temp_eventEndDate < EventEndDate
                        Need_To_Save_List.append(item)
                if Need_To_Save_List:
                    # print(EventStartDate, EventEndDate)
                    save_txt_to_disk(CalendarPath, EventStartDate + EventTextName, Need_To_Save_List)

    except Exception as ex:
        logger.error("Call method SaveEveryDayCalendarDataUsetimeMin_timeMax() error!")
        raise ex


def SaveEveryDayCalendarDataUseCreateTime(paraManyDays, paraGmailList):
    '''
    :param paraManyDays: 需要获取创建日期持续多少天内的数据,如果paraManyDays=2则获取昨天创建或更新的数据,如果paraManyDays=3,则获取前天和昨天创建或更新的数据,paraManyDays>1才有意义
    :param paraGmailList: 邮箱列表
    :return: 无
    '''
    try:
        if paraManyDays and paraGmailList:
            for x in range(int(paraManyDays), 1, -1):
                if x > 1:
                    TwoDaysAgoDate = datetime.datetime.now() - datetime.timedelta(days=x)  # 前天日期
                    TwoDaysAgoDate = TwoDaysAgoDate.strftime('%Y-%m-%d')
                    Yesterday = datetime.datetime.now() - datetime.timedelta(days=x - 1)  # 昨天日期
                    Yesterday = Yesterday.strftime('%Y-%m-%d')
                    Root_And_Branch_To_Txt = generateEveryDayCalendarData(paraGmailList, TwoDaysAgoDate, Yesterday)
                    save_txt_to_disk(CalendarPath, Yesterday + CreateTextName, Root_And_Branch_To_Txt)
    except Exception as ex:
        logger.error("Call method SaveEveryDayCalendarDataUseCreateTime() error!")
        raise ex



def MergeEventTimeData(paraEventDateList):
    '''
    :param paraEventDateList: 事件开始时间的列表
    :return: 
    '''
    try:
        allEventDataList = []
        AllColumnsList = [] #保存查分后全部80列的list
        SpecifiedCSVList = [] #保存差分后指定列event的list
        for everyDate in paraEventDateList:
            if os.path.exists(CalendarPath + everyDate + EventTextName + '.txt'):
                with open(CalendarPath + everyDate + EventTextName + '.txt', 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for line in lines:
                        if not line:  # 如果line是空
                            continue
                        else:
                            row_list = line.split(" ")
                            allEventDataList.append(row_list) #没经过任何处理的每天的event文本的集合

        for item in allEventDataList:
            if item[29] == '_':  # 不是全天事件,则事件的开始时间有'年月日'和'时分秒'
                eventStartDate = item[30][0:10]  # 取开始时间的年月日
                eventStartTime = item[30][11:19]  # 取开始时间的时分秒
            else:  # 是全天事件,则事件的开始时间只有年月日
                eventStartDate = item[29][0:10]  # 取开始时间的年月日
                eventStartTime = ''

            if item[32] == '_':  # 不是全天事件,则事件的结束时间有'年月日'和'时分秒'
                eventEndDate = item[33][0:10]  # 取结束时间的年月日
                eventEndtTime = item[33][11:19]  # 取结束时间的时分秒
            else:  # 是全天事件,则事件的结束时间只有年月日
                eventEndDate = item[32][0:10]  # 取开始时间的年月日
                eventEndtTime = ''


            if not AllColumnsList: #如果存放全部数据的List是空
                if item[13] != 'cancelled':
                    SpecifiedCSVList.append([removeBlank(item[79]), item[17], eventStartDate, eventStartTime, eventEndDate, eventEndtTime, item[42]])
                    AllColumnsList.append(item)
            else:  # 如果存放全部数据的List不是空
                idList = [x[12] for x in AllColumnsList]
                if item[13] != 'cancelled':  # 不是删除的事件
                    # iCalUID有可能会重复,所以用id,这样可以避免主键冲突
                    if item[12] not in idList:  # 没有这个id
                        SpecifiedCSVList.append([removeBlank(item[79]), item[17], eventStartDate, eventStartTime, eventEndDate, eventEndtTime, item[42]])
                        AllColumnsList.append(item)
                    else:  # 有这个id(iCalUID)
                        findIndex = idList.index(item[12])
                        # 删除原来的id对应的数据,插入该id对应的新数据(status='tentative' or status='confirmed')
                        del AllColumnsList[findIndex]
                        del SpecifiedCSVList[findIndex]
                        SpecifiedCSVList.append([removeBlank(item[79]), item[17], eventStartDate, eventStartTime, eventEndDate, eventEndtTime, item[42]])
                        AllColumnsList.append(item)
                else:  # 是删除的事件
                    if item[12] in idList:
                        delfindIndex = idList.index(item[12])
                        # 删除原来的id(iCalUID)对应的数据
                        del AllColumnsList[delfindIndex]
                        del SpecifiedCSVList[findIndex]

        # 保存全部80列差分后的events数据到CSV
        AllColumnsList = [item[:-1] + [removeBlank(item[-1])] for item in AllColumnsList] #处理最后一列的换行符
        save_data_to_csv(CalendarPath, allEvents, AllColumnsList)

        #保存指定列的数据到CSV
        with open(CalendarPath + paraEventDateList[0] + ' to ' + paraEventDateList[-2] + '_SpecifiedColumns.csv', "w", newline='',encoding="utf-8") as fo:
            title = [['calendarId', 'summary', 'eventStartDate', 'eventStartTime', 'eventEndDate', 'eventEndtTime', 'visibility']]
            writer = csv.writer(fo)
            writer.writerows(title)
            writer.writerows(SpecifiedCSVList)

        #保存全部80列到txt,这部分的数据是查分后的
        save_txt_to_disk(CalendarPath, allEvents, AllColumnsList)


    except Exception as ex:
        logger.error("Call method MergeEventTimeData() error!")
        raise ex


def MergeCreateTimeData(paraHowManyDays):
    '''
    使用allEvents.txt和每一天的create.txt合并成一个summary.txt文件
    :param paraHowManyDays: 需要获取创建日期持续多少天内的数据
    :return: 
    '''
    try:
        paraHowManyDays = int(paraHowManyDays)
        if paraHowManyDays>1:
            CreateDateList = [] #事件创建或者更新时间的List
            SummaryDataList = [] #全部数据的List,即allEvents.txt和每天的create.txt的合集
            for x in range(paraHowManyDays, 1, -1):
                if x > 1:
                    TwoDaysAgoDate = datetime.datetime.now() - datetime.timedelta(days=x)  # 前天日期
                    TwoDaysAgoDate = TwoDaysAgoDate.strftime('%Y-%m-%d')
                    Yesterday = datetime.datetime.now() - datetime.timedelta(days=x - 1)  # 昨天日期
                    Yesterday = Yesterday.strftime('%Y-%m-%d')
                    CreateDateList.append(Yesterday)

            #如果存在summary.txt,则把summary.txt中的所有内容放到SummaryDataList中
            if os.path.exists(CalendarPath + summary + '.txt'):
                with open(CalendarPath + summary + '.txt', 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for line in lines:
                        if not line:  # 如果line是空
                            continue
                        else:
                            row_list = line.split(" ")
                            SummaryDataList.append(row_list)

            # 如果存在allEvents.txt,不存在summary.txt,则把allEvents.txt中的所有内容放到SummaryDataList中
            elif os.path.exists(CalendarPath + allEvents + '.txt'):
                with open(CalendarPath + allEvents + '.txt', 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for line in lines:
                        if not line:  # 如果line是空
                            continue
                        else:
                            row_list = line.split(" ")
                            SummaryDataList.append(row_list)

            #如果事件创建或者更新时间的List不是空
            if CreateDateList:
                for everyDate in CreateDateList:
                    if os.path.exists(CalendarPath + everyDate + CreateTextName + '.txt'):
                        currentDataList = []  # 存放当前遍历的数据List
                        with open(CalendarPath + everyDate + CreateTextName + '.txt', 'r', encoding='utf-8') as f:
                            lines = f.readlines()
                            for line in lines:
                                if not line:  # 如果line是空
                                    continue
                                else:
                                    row_list = line.split(" ")
                                    currentDataList.append(row_list)
                        for row in currentDataList:
                            if not SummaryDataList:  # 如果存放全部数据的List是空
                                if row[13] != 'cancelled':
                                    SummaryDataList.append(row)
                            else:  # 如果存放全部数据的List不是空
                                # iCalUIDList = [x[43] for x in SummaryDataList]
                                idList = [x[12] for x in SummaryDataList]
                                if row[13] != 'cancelled':  # 不是删除的事件
                                    #iCalUID有可能会重复,所以用id,这样可以避免主键冲突
                                    if row[12] not in idList:  #没有这个id
                                        SummaryDataList.append(row)
                                    else:  # 有这个id(iCalUID)
                                        findIndex = idList.index(row[12])
                                        # 删除原来的id对应的数据,插入该id对应的新数据(status='tentative' or status='confirmed')
                                        del SummaryDataList[findIndex]
                                        SummaryDataList.append(row)
                                else:  # 是删除的事件
                                    if row[12] in idList:
                                        delfindIndex = idList.index(row[12])
                                        # 删除原来的id(iCalUID)对应的数据
                                        del SummaryDataList[delfindIndex]
            SummaryDataList = [item[:-1] + [removeBlank(item[-1])] for item in SummaryDataList]
            save_txt_to_disk(CalendarPath, summary, SummaryDataList)

    except Exception as ex:
        logger.error("Call method MergeCreateTimeData() error!")
        raise ex




def read_dateConfig_file_set_parameter():
    '''
    读dateConfig.ini,获取事件的开始时间、事件的结束时间、创建或者更新时间在多少天内的数据(最多29天)、要保存的文件夹的路径
    '''
    global CalendarIdListFile
    global TokenFile
    global CalendarAdministratorFile
    global SCOPES
    global CalendarPath
    global EventTimeMin
    global EventTimeMax
    global HowManyDays
    global EventTextName
    global CreateTextName
    global allEvents
    global summary
    global TokyoStartTime
    global UTCStartTime
    global UTCEndTime


    if os.path.exists(os.path.join(os.path.dirname(__file__), "dateConfig.ini")):
        try:
            conf = configparser.ConfigParser()
            conf.read(os.path.join(os.path.dirname(__file__), "dateConfig.ini"), encoding="utf-8-sig")
            CalendarIdListFile = conf.get("CalendarIdListFile", "CalendarIdListFile")
            TokenFile = conf.get("TokenFile", "TokenFile")
            CalendarAdministratorFile = conf.get("CalendarAdministratorFile", "CalendarAdministratorFile")
            SCOPES = conf.get("SCOPES", "SCOPES")
            CalendarPath = conf.get("CalendarPath", "CalendarPath")
            EventTimeMin = conf.get("EventTimeMin", "EventTimeMin")
            EventTimeMax = conf.get("EventTimeMax", "EventTimeMax")
            HowManyDays = conf.get("HowManyDays", "HowManyDays")
            EventTextName = conf.get("EventTextName", "EventTextName")
            CreateTextName = conf.get("CreateTextName", "CreateTextName")
            allEvents = conf.get("allEvents", "allEvents")
            summary = conf.get("summary", "summary")
            TokyoStartTime = conf.get("TokyoStartTime", "TokyoStartTime")
            UTCStartTime = conf.get("UTCStartTime", "UTCStartTime")
            UTCEndTime = conf.get("UTCEndTime", "UTCEndTime")
        except Exception as ex:
            logger.error("Content in dateConfig.ini has error.")
            logger.error("Exception:" + str(ex))
            raise ex
    else:
        logger.error("DateConfig.ini doesn't exist!")



if __name__ == '__main__':
    logger = write_log()  # 获取日志对象
    time_start = datetime.datetime.now()
    start = time.time()
    logger.info("Program start,now time is:" + str(time_start))
    read_dateConfig_file_set_parameter()
    gmailList = getcalendarIdList()
    Event_Date_List = getEveryDay(EventTimeMin, EventTimeMax)  #获取事件的开始到结束时间的List
    SaveEveryDayCalendarDataUsetimeMin_timeMax(Event_Date_List, gmailList) #保存事件的开始和结束时间对应的数据
    SaveEveryDayCalendarDataUseCreateTime(HowManyDays, gmailList) #保存创建事件的时间在某段范围内的数据
    MergeEventTimeData(Event_Date_List) #先合并事件开始时间和结束时间对应的数据
    # 保存差分数据
    MergeCreateTimeData(HowManyDays) #再合并创建时间(最多29天)在某段时间内的数据
    time_end = datetime.datetime.now()
    end = time.time()
    logger.info("Program end,now time is:" + str(time_end))
    logger.info("Program run : %f seconds" % (end - start))

