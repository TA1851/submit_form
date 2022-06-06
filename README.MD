---
Submit form macro

---
### Overview

#### main/sample.xlsm

UserForm に各Labelに値を入力し、入力した値をセルに書き込む。

<img src="https://github.com/TA1851/submit_form/blob/main/img/Proceedure.png">

* UserForm
  
  * Date
  * 件名
  * 作業工数(Is)
  * 作業工数(Was)
  * ww件数
  * 工数（人）

* Button

  * Clear -> 入力した値の削除
  * 登録  -> 入力した値をセルに一括入力する  
  * 閉じる -> Form を閉じる
  
<img src="https://github.com/TA1851/submit_form/blob/main/img/form.PNG">
  
---
#### Directory Configuration

<img src="https://github.com/TA1851/submit_form/blob/main/img/directory.PNG">

---
#### Requirement

* Library list

  * Visual Basic For Applicasions
  * Microsoft Excel 16.0 Object Library
  * OLE Automation
  * Microsoft Office 16.0 Object Library
  * Microsoft Forms 2.0 Object Library

---

#### Module configuration

* Form -> UserForm1

  * ExInputCls -> textbox instans create
  * Button1_Click -> Input text（登録）
  * Button2_Click -> Unload UserForm1（閉じる）
  * Button3_Click -> ExInputCls（Clear）

* 標準モジュール -> OpenForm

  * test -> UserForm1.Show vbModeless（UserFormの表示）

---
Note

<img src="https://github.com/TA1851/submit_form/blob/main/img/submit1.PNG">
<img src="https://github.com/TA1851/submit_form/blob/main/img/submit2.PNG">

---
Source code backup 取得は、VBCAを採用
[参考記事](https://tonari-it.com/vba-vbac-git/)

GitHubにPushした際に、bin folderが表示されなかった為、Pushする時に、foler nameをbin > mainに変更してPushした。
