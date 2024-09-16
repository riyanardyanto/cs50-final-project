import os
import pandas as pd
import datetime
from project import (
    Path,
    Model_bde,
    Model_cm,
    resource_path,
    get_dataframe_values,
    get_dataframe,
)

data = [
    "2024-08-15",
    "SPM",
    "LU27-ID01",
    "F5",
    "Blank stack retainer not open",
    "DH sensor stopper",
    "Valve",
    "1. mesin running normal tiba2 muncul pesan warning reduce blank hoopper empty\n2. cek kondisi hopper blank dgn kondisi kurang  dan retainer tidak bekerja \n3. EO memanggil elektrik\n4. Elektrik cek sequence secara program apakah semua syarat terpenuhi \n5. Semua syarat terpenuhi utk trigger retainer tetapi retainer tidak bekerja\n6. Cek valve dan jog manual ( kadang bisa kadang tidak ) \n7. Cleaning Valve retainer ( gerakan retainer normal )\n8. Ketika mau running kembali problem glue jet not position ( cylinder glue jet tidak bekerja )\n9. manual jog ( tidak ada gerakan utk glue jet ) proses ganti dan pinjam 18 \n10. mesin jalan kembali bertahan 30 menit ( retainer tidak membuka kembali )\n11. Cek sequence program ( trigger utk retainer tidak on karena posisi pusher belum tepat )\n12. di temukan warning clutch linear pusher ( putar manual conveyor )\n13. cek kondisi sensor posisi linear pusher ( posisi depan sensor tidak on B50.3 )\n14. Geser posisi sensor dan mesin running kembali \n15. retainer kembali problem ( di temukan sensor defect dan ganti )\n16. mesin running kembali dan masih ada problem retainer ( setting pusher titik mati depan maju ) ",
    "ganti Valve | arif | 17/08/2024\nETL dan TSG Blank Retainer | arif | 17/08/2024\nganti sensor dan propose CL sensor 50.2 | kama | 17/08/2024\n Propose MP&S pengecekan CL sensor 50.2  | kama | 17/08/2024\npropose MP&S pengecekan backlast linear pusher | Riyan | 17/08/2024\nFix Plat pembacaan sensor (tambah lubang baut) | Riyan | 17/08/2024",
    "Valves / Pumps",
    "VMPA1-M1H-M-PI / 533342 / 12841507",
    "137",
    "60",
    "Incorrect Intallation / Termination/ Kesalahan dalam melakukan installation",
]

extraction = {
    "Date": {0: datetime.date(2024, 8, 15)},
    "Dept": {0: "SPM"},
    "LU": {0: "LU27-ID01"},
    "Technology": {0: "F5"},
    "Loss Name": {0: "Blank stack retainer not open"},
    "Evidence": {0: "DH sensor stopper"},
    "Component/ Part": {0: "Valve"},
    "Chronology": {
        0: "1. mesin running normal tiba2 muncul pesan warning reduce blank hoopper empty\n2. cek kondisi hopper blank dgn kondisi kurang  dan retainer tidak bekerja \n3. EO memanggil elektrik\n4. Elektrik cek sequence secara program apakah semua syarat terpenuhi \n5. Semua syarat terpenuhi utk trigger retainer tetapi retainer tidak bekerja\n6. Cek valve dan jog manual ( kadang bisa kadang tidak ) \n7. Cleaning Valve retainer ( gerakan retainer normal )\n8. Ketika mau running kembali problem glue jet not position ( cylinder glue jet tidak bekerja )\n9. manual jog ( tidak ada gerakan utk glue jet ) proses ganti dan pinjam 18 \n10. mesin jalan kembali bertahan 30 menit ( retainer tidak membuka kembali )\n11. Cek sequence program ( trigger utk retainer tidak on karena posisi pusher belum tepat )\n12. di temukan warning clutch linear pusher ( putar manual conveyor )\n13. cek kondisi sensor posisi linear pusher ( posisi depan sensor tidak on B50.3 )\n14. Geser posisi sensor dan mesin running kembali \n15. retainer kembali problem ( di temukan sensor defect dan ganti )\n16. mesin running kembali dan masih ada problem retainer ( setting pusher titik mati depan maju ) "
    },
    "Countermeasure": {
        0: "ganti Valve | arif | 17/08/2024\nETL dan TSG Blank Retainer | arif | 17/08/2024\nganti sensor dan propose CL sensor 50.2 | kama | 17/08/2024\n Propose MP&S pengecekan CL sensor 50.2  | kama | 17/08/2024\npropose MP&S pengecekan backlast linear pusher | Riyan | 17/08/2024\nFix Plat pembacaan sensor (tambah lubang baut) | Riyan | 17/08/2024"
    },
    "Object Part": {0: "Valves / Pumps"},
    "SAP#/OEM part number#: ": {0: "VMPA1-M1H-M-PI / 533342 / 12841507"},
    "Down time:": {0: "137"},
    "Repair time:": {0: "60"},
    "ACC  Accident or damage": {
        0: "Incorrect Intallation / Termination/ Kesalahan dalam melakukan installation"
    },
}


def test_resource_path():
    assert resource_path("bde_image.jpg") == os.path.join(
        os.path.abspath("."), "bde_image.jpg"
    )


def test_get_dataframe_values():
    db_path = Path("./Database/Database_BDE.db")
    model_bde = Model_bde(db_path)
    model_cm = Model_cm(db_path)

    assert len(get_dataframe_values(data, model_bde, model_cm)) == 2
    assert type(get_dataframe_values(data, model_bde, model_cm)[1]) == list
    assert get_dataframe_values(data, model_bde, model_cm)[1][0][3] == "SPM"


def test_get_dataframe():
    assert get_dataframe(data).equals(pd.DataFrame(extraction))
