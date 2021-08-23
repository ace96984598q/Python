# encoding = utf8
import datetime
import smtplib
import cx_Oracle
import os
import pandas as pd
import configparser
from sqlalchemy import create_engine
from email.mime.text import MIMEText  # 用於製作文字內文
from email.mime.multipart import MIMEMultipart  # 用於建立郵件載體
from email.mime.base import MIMEBase  # 用於承載附檔
from email import encoders  # 用於附檔編碼

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.AL32UTF8'

cf = configparser.ConfigParser()
cf.read("config.ini")  # 需於config.ini中建立[config_info]並設定參數


def get_goods_info():
    host = cf.get("config_info", "host")
    port = cf.get("config_info", "port")
    sid = cf.get("config_info", "sid")
    user = cf.get("config_info", "user")
    password = cf.get("config_info", "password")
    sid = cx_Oracle.makedsn(host, port, sid=sid)
    db = 'oracle://{user}:{password}@{sid}'.format(user=user, password=password, sid=sid)
    engine = create_engine(db, pool_recycle=10, pool_size=50, max_identifier_length=128,
                           echo=False)
    conn = engine.connect()
    sql = "SELECT INSP_DT.COMP_COD ,WHOUSE_RF.WHOUSE_NAM ,GOODS_MN.BGROUP_COD ,INSP_DT.GOODS_COD ,GOODS_MN.GOODS_SNA " \
          ",UNIT_RF.UNIT_NAM ,SUBSTR(TO_CHAR(INSP_MN.INSP_DAT,'yyyy/mm/dd'),0,10) ,INSP_MN.INSP_NOS ," \
          "INSP_MN.PURCH_NOS ,SUPLY_MN.SUPLY_SNA ,INSP_DT.INSP_QNT,INSP_DT.INSP_AMT ,INSP_DT.SINSP_AMT, " \
          "PURCH_MN.PURMAN_COD FROM INSP_DT LEFT JOIN GOODS_MN ON GOODS_MN.COMP_COD = INSP_DT.COMP_COD AND " \
          "GOODS_MN.SYS_FLAG = INSP_DT.SYS_FLAG AND GOODS_MN.GOODS_COD = INSP_DT.GOODS_COD LEFT JOIN INSP_MN ON " \
          "INSP_DT.COMP_COD = INSP_MN.COMP_COD AND INSP_DT.SYS_FLAG = INSP_MN.SYS_FLAG AND INSP_DT.INSP_NOS = " \
          "INSP_MN.INSP_NOS LEFT JOIN UNIT_RF ON UNIT_RF.UNIT_TYP = INSP_DT.PUNIT_TYP AND UNIT_RF.COMP_COD = " \
          "INSP_DT.COMP_COD LEFT JOIN WHOUSE_RF ON WHOUSE_RF.COMP_COD = INSP_DT.COMP_COD AND WHOUSE_RF.SYS_FLAG = " \
          "INSP_DT.SYS_FLAG AND WHOUSE_RF.WHOUSE_COD = INSP_DT.WHOUSE_COD LEFT JOIN SUPLY_MN ON SUPLY_MN.SUPLY_COD = " \
          "INSP_MN.SUPLY_COD AND SUPLY_MN.COMP_COD = INSP_DT.COMP_COD LEFT JOIN PURCH_MN ON PURCH_MN.PURCH_NOS = " \
          "INSP_MN.PURCH_NOS AND PURCH_MN.COMP_COD =INSP_DT.COMP_COD WHERE SUBSTR(TO_CHAR(INSP_MN.INSP_DAT," \
          "'yyyymmdd'),0,10) BETWEEN TO_CHAR(SYSDATE+(2-TO_CHAR(SYSDATE,'d'))-6,'yyyymmdd') and to_char(sysdate," \
          "'yyyymmdd') ORDER BY COMP_COD,SUBSTR(TO_CHAR(INSP_MN.INSP_DAT,'yyyy/mm/dd'),0,10) "
    frame = pd.read_sql(sql=sql, con=engine)
    frame_newname = frame.rename(
        columns={"comp_cod": "館別", "whouse_nam": "進貨倉庫", "bgroup_cod": "類別", "goods_cod": "貨品代號",
                 "goods_sna": "貨品描述(規格)", "unit_nam": "單位",
                 "SUBSTR(TO_CHAR(INSP_MN.INSP_DAT,'YYYY/MM/DD'),0,10)": "驗收日期", "insp_nos": "驗收單號", "purch_nos": "採購單號",
                 "suply_sna": "廠商", "insp_qnt": "驗收量", "insp_amt": "驗收單價", "sinsp_amt": "驗收總成本", "purman_cod": "採購人員"})
    frame_newname.to_excel(get_today() + '.xlsx', index=False)
    print("檔案產生成功")
    conn.close()


def get_today():
    today = datetime.date.today()
    today_right = today.strftime("%Y%m%d")
    return today_right


def get_yesterday():
    today = datetime.date.today()
    one_day = datetime.timedelta(days=1)
    yesterday = today - one_day
    date_right = yesterday.strftime("%Y%m%d")
    return date_right


def send_email(subject: str, body: str, name: str):
    get_goods_info()
    mail_from = cf.get("mail", "email")
    mail_pass = cf.get("mail", "password")
    email_list = cf.get("mail", "send_list").split(',')  # split用於將[]中逗號左右側資料掛上''
    service_email = ','.join(email_list)
    mime = MIMEMultipart()
    mime["Subject"] = subject
    mime["To"] = service_email
    mime["From"] = name
    mime.attach(MIMEText(body))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(get_today() + ".xlsx", "rb").read())  # 開啟xlsx且讀取
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename= "Weekly Report.xlsx"')
    mime.attach(part)
    msg = mime.as_string()
    smtp = smtplib.SMTP("smtp.gmail.com", 587)  # gmail郵件port
    smtp.ehlo()  # 向smtp伺服器標示自身
    smtp.starttls()  # tls連線
    smtp.login(mail_from, mail_pass)  # login帳號
    from_addr = mail_from  # 寄信來源歸屬
    to_addr = email_list  # service_email
    status = smtp.sendmail(from_addr, to_addr, msg)
    if status == {}:
        print("郵件傳送成功!")
    else:
        print("郵件傳送失敗!")
    smtp.quit()


if __name__ == '__main__':
    send_email('採購週報表(時間區間為上周二至今日)', '請查閱附件', '系統自動發送')
