import os
import cx_Oracle
import pandas as pd

uploadQuery = "insert into TBL_KYCPHOTO_INPUT (CUST_ID,CUST_PHOTO,KYC_PHOTO)values (:custID,:custPhoto,:kycPhoto)"

custData = pd.read_excel("IMAGES EXCEL.xlsx")

custData['query'] = custData.apply(lambda x: str(x['GLB ID']).replace("GLBCUST",""),axis = 1)

custData = custData[['CUST_ID','query']]


#custImg = ""
def get_img(query):

    try:
        with open(f"C:/crf/ckyc/OneDrive_2022-09-15/LOT 6/{query[1]}_1.jpg","rb") as f:
            custImg = f.read()

        with open(f"C:/crf/ckyc/OneDrive_2022-09-15/LOT 6/{query[1]}_2.jpg","rb") as f:
            kycPhoto = f.read()

        cur.execute(f"{uploadQuery}",custID = query[0],custPhoto = custImg ,kycPhoto = kycPhoto)
        #con.commit()
        print(f"Uploading : {query[0]} . . .")
        return True
    except cx_Oracle.IntegrityError:
        print(f"failed to added due to IntegrityError .")
        return False
    except cx_Oracle.DatabaseError:
        print(f"failed to added due to DatabaseError .")
        return False
    except Exception as err:
        print(err)
        return False

with cx_Oracle.connect("username", "password", "dbname") as con:
    cur = con.cursor()
    custData['status'] = custData.apply(lambda x: get_img(x),axis = 1)

custData.to_excel("./upload_status.xlsx")
