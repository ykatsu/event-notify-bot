###
Google Calendar のイベント通知をSlackに投稿します
GASのトリガーからの１分毎の起動を想定しています。
１分毎なので、つまりは１分程度の誤差は出るという事です。(仮にサーバーがキッチリ動いたとしても)
終日予定の通知の取得方法が不明な為、それに関しては通知されません。
###

# SlackAppライブラリキー:M3W5Ut3Q39AaIwLquryEPMwV62A3znfOO
token = '<<自分のSlackのtoken>>' # Slackのtoken
channel_id = '<<投稿するチャンネルのid>>'  # チャンネルID https://api.slack.com/methods/channels.list/testで調べられる
calendar_id = '<<カレンダーid>>'
bot_name = 'notify'
bot_icon = ':robot_face:'

isDebug = true

###
Post処理メイン
###
doPost = ->
  now = new Date()
  endTime = new Date(now)
  endTime.setDate(endTime.getDate() + 7) # 1週間先までの予定の通知を調べます
  message = listupEventNotify(calendar_id, now, endTime)
  Logger.log message
  postSlack(message)
  return

###
通知をリストアップ
###
listupEventNotify = (cal_id, now, endTime) ->
  # 前回処理した時間を取得し、現在の時間を保存
  prevTime = getPrevTmve()
  nt = fd(now)
  setCurTime(nt)

  list = []
  events = CalendarApp.getCalendarById(cal_id).getEvents(now, endTime)
  if isDebug
    n = fd(now)
    e = fd(endTime)
    Logger.log "now = #{n}, end = #{e}"
    Logger.log 'Number of events: ' + events.length

  for event in events
    reminders = event.getPopupReminders() # 終日予定はなぜかこれが取れない
    remindHits = []
    for remindMinutesTo in reminders
      eventTime = event.getStartTime()
      remindTime = new Date(eventTime)
      remindTime.setTime(eventTime.getTime() - remindMinutesTo*60000)

      if isDebug
        rt = fd(remindTime)
        pt = fd(prevTime)
        Logger.log "prev:#{pt}, remind:#{rt}, now:#{nt}"

      if prevTime < remindTime <= now
        remindHits.push(remindMinutesTo)

    if remindHits.length >= 1
      # 通知が複数該当する場合は、直近のものだけ通知
      minutesTo = Math.min.apply(null, remindHits)
      Logger.log "minhit:#{minutesTo} in #{remindHits.join(',')}" if isDebug
      list.push(generateNotifyMessage(event.getTitle(), minutesTo, eventTime))

  text = ''
  for l in list
    text += l + '\n'
  text

###
日付フォーマット
###
fd = (time) ->
  Utilities.formatDate(time, 'GMT+0900', 'yyyy/MM/dd HH:mm:ss')

###
シート操作関連
###
setCurTime = (nt) ->
  getSheet().getRange(1,2,1,1).setValue(nt)

getPrevTmve = ->
  prev = getSheet().getRange(1,2,1,1).getValue()
  if prev == ''
    prev = new Date()
    prev.setDate(prev.getDate() - 1)
  prev

getSheet = ->
  if getSheet.sheet
    return getSheet.sheet
  getSheet.sheet = SpreadsheetApp.getActive().getSheetByName('シート1');

###
通知イベント表示用メッセージを作成
###
generateNotifyMessage = (title, minutesTo, eventTime) ->
  remindTime = new Date(eventTime.getTime() - minutesTo*60000)
  dayDiff = calcDayDiff(eventTime, remindTime)
  if minutesTo <= 0
    mes = "#{title}の時間です。"
  else if (dayDiff == 1)
    mes = "明日は、#{title}です。"
  else if (dayDiff >= 2)
    mes = "#{title}の#{dayDiff}日前です。"
  else if minutesTo < 60
    mes = "#{title}の#{minutesTo}分前です。"
  else
    mes = "#{title}の#{Math.round(minutesTo/60)}時間前です。"
  mes

calcDayDiff = (eventTime, remindTime) ->
  eventDay  = (eventTime.getTime()  + (1000 * 60 * 60 * 9))/(1000 * 60 * 60 * 24) >> 0
  remindDay = (remindTime.getTime() + (1000 * 60 * 60 * 9))/(1000 * 60 * 60 * 24) >> 0
  Logger.log "eventDay:#{eventDay}, remindDay:#{remindDay}" if isDebug
  eventDay - remindDay

###
Slackへポスト
###
postSlack = (payload) ->
  if payload.trim() == ''
    return

  app = SlackApp.create(token)
  app.postMessage(channel_id, payload,
    username: bot_name
    icon_emoji: bot_icon)
  return

###
# テスト用メソッド
###
testGenerateNotifyMessage = ->
  eventTime = new Date(Date.parse("2015/11/27 23:59:00"))

  Logger.log generateNotifyMessage("イベント", -1, eventTime)
  Logger.log generateNotifyMessage("イベント", 0, eventTime)
  Logger.log generateNotifyMessage("イベント", 59, eventTime)
  Logger.log generateNotifyMessage("イベント", 60, eventTime)
  Logger.log generateNotifyMessage("イベント", 61, eventTime)
  Logger.log generateNotifyMessage("イベント", 80, eventTime)
  Logger.log generateNotifyMessage("イベント", 100, eventTime)
  Logger.log generateNotifyMessage("イベント", 60*2, eventTime)
  Logger.log generateNotifyMessage("イベント", 60*23, eventTime)
  Logger.log generateNotifyMessage("イベント", 60*24, eventTime)
  Logger.log generateNotifyMessage("イベント", 60*24+60*12, eventTime)
  Logger.log generateNotifyMessage("イベント", 60*24*2, eventTime)
  Logger.log generateNotifyMessage("イベント", 60*24*7, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 12:00:00"))
  Logger.log generateNotifyMessage("イベント", 60*24, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 12:00:00"))
  Logger.log generateNotifyMessage("イベント", 60*24*2, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 12:00:00"))
  Logger.log generateNotifyMessage("イベント", 60*24*5, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 00:00:00"))
  Logger.log generateNotifyMessage("イベント", 1, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 01:00:00"))
  Logger.log generateNotifyMessage("イベント", 60, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 00:59:00"))
  Logger.log generateNotifyMessage("イベント", 60, eventTime)

  eventTime = new Date(Date.parse("2015/11/27 00:00:00"))
  Logger.log generateNotifyMessage("イベント", 60*25, eventTime)

  return

testListupEventNotify = ->
  if true
    now = new Date(Date.parse("2015/11/27 12:00:00"))
  else
    now = new Date()
  endTime = new Date(now)
  message = listupEventNotify(calendar_id, now, endTime)
  return

test = ->
  # testListupEventNotify()
  testGenerateNotifyMessage()
  return
