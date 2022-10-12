---
title: Outlook アイテムからエンティティ文字列を抽出する
description: Outlook アドイン内の Outlook アイテムからエンティティを抽出する方法について説明します。
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 512999272c720f5b87480c49d60c649bc6a886ad
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541276"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a>Outlook アイテムからエンティティ文字列を抽出する

This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.

> [!NOTE]
> この記事で説明する Outlook アドイン機能では、Office アドイン [の Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用するアドインではサポートされていないアクティブ化規則を使用します。

サポートされるエンティティには次のようなものがあります。

- **住所**: 米国の郵送先住所で、少なくとも番地、住所、都市名、州名、郵便番号の各要素の一部を含むもの。

- **連絡先**: 住所、勤務先名など、他のエンティティのコンテキストにおける、個人の連絡先情報。

- **電子メール アドレス**: SMTP 電子メール アドレス。

- **Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.

- **電話番号**: 北米の電話番号。

- **タスクの提案**: 通常、実行可能な語句で表現される、タスクの提案。

- **URL**

Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.

Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.

The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.

## <a name="xml-manifest"></a>XML マニフェスト

エンティティ アドインには、論理 OR 演算で結合された 2 つのアクティブ化ルールがあります。

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

これらのルールでは、閲覧ウィンドウまたは閲覧インスペクターの現在選択されているアイテムが予定またはメッセージ (電子メール メッセージ、会議出席依頼、会議出席依頼の返信、または会議の取り消しなど) であるときに、Outlook でこのアドインをアクティブ化することを指定しています。

The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```

## <a name="html-implementation"></a>HTML の実装

The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.

すべての Outlook アドインに office.js を含める必要があります。 次の HTML ファイルには、コンテンツ配信ネットワーク (CDN) 上のバージョン 1.1 のoffice.jsが含まれています。

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```

## <a name="style-sheet"></a>スタイル シート

The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.

```CSS
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```

## <a name="javascript-implementation"></a>JavaScript の実装

残りのセクションでは、このサンプル (default_entities.js ファイル) を使用して、ユーザーが表示中のメッセージまたは予定の件名と本文から一般的なエンティティを抽出する方法について説明します。

## <a name="extracting-entities-upon-initialization"></a>初期化時のエンティティの抽出

[Office.initialize](/javascript/api/office#Office_initialize_reason_) イベントが発生すると、エンティティ アドインは現在のアイテムの [getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッドを呼び出します。 このメソッドは `getEntities` 、サポートされているエンティティのインスタンスの配列をグローバル変数 `_MyEntities` に返します。 関連する JavaScript コードを次に示します。

```js
// Global variables
let _Item;
let _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    const _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}
```

## <a name="extracting-addresses"></a>住所の抽出

ユーザーが **[Get Addresses]** ボタンをクリックすると、`myGetAddresses` イベント ハンドラーが `_MyEntities` オブジェクトの [addresses](/javascript/api/outlook/office.entities#outlook-office-entities-addresses-member) プロパティから住所の配列を取得します (住所が抽出されていた場合)。 抽出された各住所は、配列内の文字列として保存されます。 `myGetAddresses` はローカル HTML 文字列を `htmlText` で生成し、抽出された住所のリストを表示します。 関連する JavaScript コードを次に示します。

```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    let htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    const addressesArray = _MyEntities.addresses;
    for (let i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-contact-information"></a>連絡先情報の抽出

ユーザーが **[連絡先情報の取得** ] ボタンをクリックすると、 `myGetContacts` イベント ハンドラーは、抽出された場合、オブジェクトの [contacts](/javascript/api/outlook/office.entities#outlook-office-entities-contacts-member) プロパティから連絡先の配列とその情報を `_MyEntities` 取得します。 抽出された各連絡先は、配列内の [Contact](/javascript/api/outlook/office.contact) オブジェクトとして格納されます。 `myGetContacts` は、各連絡先に関するその他のデータを取得します。 コンテキストによって、Outlook がアイテム&mdash;から電子メール メッセージの最後に署名を抽出できるかどうか、または連絡先の近くに次の情報の少なくとも一部が存在する必要があるかどうかを判断します。

- [Contact.personName](/javascript/api/outlook/office.contact#outlook-office-contact-personname-member) プロパティから取得される連絡先の名前を表す文字列。

- [Contact.businessName](/javascript/api/outlook/office.contact#outlook-office-contact-businessname-member) プロパティから取得される連絡先に関連付けられた会社名を表す文字列。

- The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#outlook-office-contact-phonenumbers-member) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.

- 電話番号配列内の **PhoneNumber** メンバーごとの、[PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-phonestring-member) プロパティから取得される電話番号を表す文字列。

- The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#outlook-office-contact-urls-member) property. Each URL is represented as a string in an array member.

- The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#outlook-office-contact-emailaddresses-member) property. Each email address is represented as a string in an array member.

- The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#outlook-office-contact-addresses-member) property. Each postal address is represented as a string in an array member.

`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.

```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    let htmlText = "";

    // Gets an array of contacts and their information.
    const contactsArray = _MyEntities.contacts;
    for (let i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        let phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (let j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        let urlsArray = contactsArray[i].urls;
        for (let j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        let emailAddressesArray = contactsArray[i].emailAddresses;
        for (let j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        let addressesArray = contactsArray[i].addresses;
        for (let j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-email-addresses"></a>電子メール アドレスの抽出

When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#outlook-office-entities-emailaddresses-member) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.

```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    let htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    const emailAddressesArray = _MyEntities.emailAddresses;
    for (let i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-meeting-suggestions"></a>会議提案の抽出

ユーザーが **[Get Meeting Suggestions]** ボタンをクリックすると、`myGetMeetingSuggestions` イベント ハンドラーが `_MyEntities` オブジェクトの [meetingSuggestions](/javascript/api/outlook/office.entities#outlook-office-entities-meetingsuggestions-member) プロパティから会議提案の配列を取得します (会議提案が抽出されていた場合)。

 > [!NOTE]
 > エンティティの種類は、メッセージのみがサポートされますが、予定はサポート `MeetingSuggestion` されません。

抽出された各会議提案は、[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) オブジェクトとして配列に格納されます。 は、各会議提案に関する次の詳細なデータを取得します。

- [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-meetingstring-member) プロパティから取得される会議提案として識別された文字列。

- The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-attendees-member) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.

- 出席者ごとの、[EmailUser.displayName](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-displayname-member) プロパティから取得される名前。

- 出席者ごとの、[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-emailaddress-member) プロパティから取得される SMTP アドレス。

- [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-location-member) プロパティから取得される、会議提案の場所を表す文字列。

- [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-subject-member) プロパティから取得される、会議提案の議題を表す文字列。

- [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-start-member) プロパティから取得される、会議提案の開始時刻を表す文字列。

- [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#outlook-office-meetingsuggestion-end-member) プロパティから取得される会議提案の終了時刻を表す文字列。

`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.

```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    let htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    const meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (let i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        let attendeesArray = meetingsArray[i].attendees;
        for (let j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-phone-numbers"></a>電話番号の抽出

When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#outlook-office-entities-phonenumbers-member) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:

- [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-type-member) プロパティから取得される電話番号の種類 (自宅の電話番号など) を表す文字列。

- [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-phonestring-member) プロパティから取得される、実際の電話番号を表す文字列。

- [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#outlook-office-phonenumber-originalphonestring-member) プロパティから取得される電話番号として最初に識別された文字列。

`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.

```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    let htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    const phoneNumbersArray = _MyEntities.phoneNumbers;
    for (let i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-task-suggestions"></a>タスクの提案の抽出

When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#outlook-office-entities-tasksuggestions-member) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:

- [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#outlook-office-tasksuggestion-taskstring-member) プロパティから取得されるタスクの提案として最初に識別された文字列。

- The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#outlook-office-tasksuggestion-assignees-member) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.

- 割り当て先ごとの、[EmailUser.displayName](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-displayname-member) プロパティから取得される名前。

- 割り当て先ごとの、[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#outlook-office-emailuser-emailaddress-member) プロパティから取得される SMTP アドレス。

`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.

```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    let htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    const tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (let i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        let assigneesArray = tasksArray[i].assignees;
        for (let j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="extracting-urls"></a>URL の抽出

When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#outlook-office-entities-urls-member) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.

```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    let htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    const urlArray = _MyEntities.urls;
    for (let i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="clearing-displayed-entity-strings"></a>表示されたエンティティ文字列の消去

Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.

```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```

## <a name="javascript-listing"></a>JavaScript の内容

次に、JavaScript の実装の内容全体を示します。

```js
// Global variables
let _Item;
let _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    const _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    let htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    const addressesArray = _MyEntities.addresses;
    for (let i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    let htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    const emailAddressesArray = _MyEntities.emailAddresses;
    for (let i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    let htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    const meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (let i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        let attendeesArray = meetingsArray[i].attendees;
        for (let j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    let htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    const phoneNumbersArray = _MyEntities.phoneNumbers;
    for (let i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    let htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    const tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (let i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        let assigneesArray = tasksArray[i].assignees;
        for (let j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    let htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    const urlArray = _MyEntities.urls;
    for (let i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```

## <a name="see-also"></a>関連項目

- [閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
- [item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) メソッド
