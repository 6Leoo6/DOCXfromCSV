# DOCXfromCSV
Generate multiple docx files based on an uploaded csv


A feltöltés menete:
  1. Docker Daemon elindítása
  2. Új projekt készítése az AWS ECR-en
  3. A View push commands gomb (ECR projekt oldala) megnyomása után láthatóvá vált parancsokat kimásolni egyenként sorban beilleszteni őket egy a fő mappában megnyitott CMD-be
  4. A docker image sikeres feltöltése után új lambda function készítése az alábbi beállításokkal:
      -Container Image
      -A Container URI bemásolása
  5. A léterhozott function oldalán a configuration menüpontnál létre kell hozni egy új funtion url-t és a general configuration-nél be kell álítani egy minél nagyobb mennyiségű ramot és egy min. 1 perces timeout-ot
