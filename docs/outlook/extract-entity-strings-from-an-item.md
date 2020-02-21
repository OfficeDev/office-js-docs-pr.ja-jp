---
title: Outlook アイテムからエンティティ文字列を抽出する
description: Outlook アドイン内の Outlook アイテムからエンティティを抽出する方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 95f88b6bcd47cbfd85de89a3de89d9a9e2fe571f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166505"
---
# <a name="extract-entity-strings-from-an-outlook-item"></a><span data-ttu-id="5f22f-103">Outlook アイテムからエンティティ文字列を抽出する</span><span class="sxs-lookup"><span data-stu-id="5f22f-103">Extract entity strings from an Outlook item</span></span>

<span data-ttu-id="5f22f-p101">この記事では、選択した Outlook アイテムの件名と本文に含まれる、サポートされる既知のエンティティの文字列インスタンスを抽出する **[エンティティの表示]** Outlook アドインを作成する方法について説明します。対象のアイテムは、予定、メール メッセージ、会議出席依頼、会議出席依頼の返信、または会議の取り消しです。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p101">This article describes how to create a **Display entities** Outlook add-in that extracts string instances of supported well-known entities in the subject and body of the selected Outlook item. This item can be an appointment, email message, or meeting request, response, or cancellation.</span></span> 

<span data-ttu-id="5f22f-106">サポートされるエンティティには次のようなものがあります。</span><span class="sxs-lookup"><span data-stu-id="5f22f-106">The supported entities include:</span></span>

- <span data-ttu-id="5f22f-107">**住所**: 米国の郵送先住所で、少なくとも番地、住所、都市名、州名、郵便番号の各要素の一部を含むもの。</span><span class="sxs-lookup"><span data-stu-id="5f22f-107">**Address**: A United States postal address, that has at least a subset of the elements of a street number, street name, city, state, and zip code.</span></span>
    
- <span data-ttu-id="5f22f-108">**連絡先**: 住所、勤務先名など、他のエンティティのコンテキストにおける、個人の連絡先情報。</span><span class="sxs-lookup"><span data-stu-id="5f22f-108">**Contact**: A person's contact information, in the context of other entities such as an address or business name.</span></span>
    
- <span data-ttu-id="5f22f-109">**電子メール アドレス**: SMTP 電子メール アドレス。</span><span class="sxs-lookup"><span data-stu-id="5f22f-109">**Email address**: An SMTP email address.</span></span>
    
- <span data-ttu-id="5f22f-p102">**会議提案**: イベントへの言及などの会議提案。予定ではなくメッセージのみが会議提案の抽出をサポートすることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p102">**Meeting suggestion**: A meeting suggestion, such as a reference to an event. Note that only messages but not appointments support extracting meeting suggestions.</span></span>
    
- <span data-ttu-id="5f22f-112">**電話番号**: 北米の電話番号。</span><span class="sxs-lookup"><span data-stu-id="5f22f-112">**Phone number**: A North American phone number.</span></span>
    
- <span data-ttu-id="5f22f-113">**タスクの提案**: 通常、実行可能な語句で表現される、タスクの提案。</span><span class="sxs-lookup"><span data-stu-id="5f22f-113">**Task suggestion**: A task suggestion, typically expressed in an actionable phrase.</span></span>
    
- <span data-ttu-id="5f22f-114">**URL**</span><span class="sxs-lookup"><span data-stu-id="5f22f-114">**URL**</span></span>
    
<span data-ttu-id="5f22f-p103">これらのエンティティの大部分は、大量のデータの機械学習に基づいた自然言語認識を利用しています。このため、認識は非確定的で、結果が Outlook アイテムの特定のコンテキストに左右されることがあります。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p103">Most of these entities rely on natural language recognition, which is based on machine learning of large amounts of data. This recognition is nondeterministic and sometimes depends on the context in the Outlook item.</span></span>

<span data-ttu-id="5f22f-p104">ユーザーが予定、メール メッセージ、会議出席依頼、会議出席依頼の返信、会議の取り消しの表示を選択するたびに、Outlook によってエンティティ アドインがアクティブ化されます。初期化時に、このサンプルのエンティティ アドインは、現在のアイテムからサポートされているエンティティのすべてのインスタンスを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p104">Outlook activates the entities add-in whenever the user selects an appointment, email message, or meeting request, response, or cancellation for viewing. During initialization, the sample entities add-in reads all instances of the supported entities from the current item.</span></span> 

<span data-ttu-id="5f22f-p105">このアドインにはユーザーがエンティティの種類を選択するためのボタンがあります。ユーザーがエンティティを選択すると、アドインは選択されたエンティティのインスタンスをアドイン ウィンドウに表示します。後続の各セクションでは、エンティティ アドインの XML マニフェスト、HTML ファイル、および JavaScript ファイルの内容を示し、それぞれのエンティティの抽出をサポートするコードについて詳しく説明します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p105">The add-in provides buttons for the user to choose a type of entity. When the user selects an entity, the add-in displays instances of the selected entity in the add-in pane. The following sections list the XML manifest, and HTML and JavaScript files of the entities add-in, and highlight the code that supports the respective entity extraction.</span></span>

## <a name="xml-manifest"></a><span data-ttu-id="5f22f-122">XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="5f22f-122">XML manifest</span></span>

<span data-ttu-id="5f22f-123">エンティティ アドインには、論理 OR 演算で結合された 2 つのアクティブ化ルールがあります。</span><span class="sxs-lookup"><span data-stu-id="5f22f-123">The entities add-in has two activation rules joined by a logical OR operation.</span></span> 

```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

<span data-ttu-id="5f22f-124">これらのルールでは、閲覧ウィンドウまたは閲覧インスペクターの現在選択されているアイテムが予定またはメッセージ (電子メール メッセージ、会議出席依頼、会議出席依頼の返信、または会議の取り消しなど) であるときに、Outlook でこのアドインをアクティブ化することを指定しています。</span><span class="sxs-lookup"><span data-stu-id="5f22f-124">These rules specify that Outlook should activate this add-in when the currently selected item in the Reading Pane or read inspector is an appointment or message (including an email message, or meeting request, response, or cancellation).</span></span>

<span data-ttu-id="5f22f-p106">エンティティ アドインのマニフェストを次に示します。このマニフェストは、Office アドイン マニフェストのスキーマ バージョン 1.1 を使用します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p106">The following is the manifest of the entities add-in. It uses version 1.1 of the schema for Office Add-ins manifests.</span></span>

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


## <a name="html-implementation"></a><span data-ttu-id="5f22f-127">HTML の実装</span><span class="sxs-lookup"><span data-stu-id="5f22f-127">HTML implementation</span></span>

<span data-ttu-id="5f22f-p107">エンティティ アドインの HTML ファイルでは、ユーザーがエンティティの種類を選択するためのボタンと、表示されたエンティティのインスタンスを消去するためのボタンを指定しています。このファイルでは、後の「[JavaScript の実装](#javascript-implementation)」で説明する default_entities.js という JavaScript ファイルを指定しています。JavaScript ファイルには、それぞれのボタンに対するイベント ハンドラーが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p107">The HTML file of the entities add-in specifies buttons for the user to select each type of entity, and another button to clear displayed instances of an entity. It includes a JavaScript file, default_entities.js, which is described in the next section under [JavaScript implementation](#javascript-implementation). The JavaScript file includes the event handlers for each of the buttons.</span></span>

<span data-ttu-id="5f22f-p108">すべての Outlook アドインに office.js を含める必要があります。以下の HTML ファイルには、CDN に office.js のバージョン 1.1 が含まれます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p108">Note that all Outlook add-ins must include office.js. The HTML file that follows includes version 1.1 of office.js on the CDN.</span></span> 

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


## <a name="style-sheet"></a><span data-ttu-id="5f22f-133">スタイル シート</span><span class="sxs-lookup"><span data-stu-id="5f22f-133">Style sheet</span></span>


<span data-ttu-id="5f22f-p109">エンティティ アドインでは、default_entities.css というオプションの CSS ファイルを使用して出力のレイアウトを指定しています。次に、この CSS ファイルの内容を示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p109">The entities add-in uses an optional CSS file, default_entities.css, to specify the layout of the output. The following is a listing of the CSS file.</span></span>


```CSS
*
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


## <a name="javascript-implementation"></a><span data-ttu-id="5f22f-136">JavaScript の実装</span><span class="sxs-lookup"><span data-stu-id="5f22f-136">JavaScript implementation</span></span>

<span data-ttu-id="5f22f-137">残りのセクションでは、このサンプル (default_entities.js ファイル) を使用して、ユーザーが表示中のメッセージまたは予定の件名と本文から一般的なエンティティを抽出する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-137">The remaining sections describe how this sample (default_entities.js file) extracts well-known entities from the subject and body of the message or appointment that the user is viewing.</span></span>

## <a name="extracting-entities-upon-initialization"></a><span data-ttu-id="5f22f-138">初期化時のエンティティの抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-138">Extracting entities upon initialization</span></span>

<span data-ttu-id="5f22f-p110">[Office.initialize](/javascript/api/office#office-initialize-reason-) イベントが発生すると、エンティティ アドインは現在のアイテムの [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッドを呼び出します。**getEntities** メソッドは、グローバル変数 `_MyEntities` を返します。この変数は、サポートされているエンティティのインスタンスの配列です。関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p110">Upon the [Office.initialize](/javascript/api/office#office-initialize-reason-) event, the entities add-in calls the [getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method of the current item. The **getEntities** method returns the global variable `_MyEntities` an array of instances of supported entities. The following is the related JavaScript code.</span></span>


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## <a name="extracting-addresses"></a><span data-ttu-id="5f22f-142">住所の抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-142">Extracting addresses</span></span>


<span data-ttu-id="5f22f-143">ユーザーが **[Get Addresses]** ボタンをクリックすると、`myGetAddresses` イベント ハンドラーが `_MyEntities` オブジェクトの [addresses](/javascript/api/outlook/office.entities#addresses) プロパティから住所の配列を取得します (住所が抽出されていた場合)。</span><span class="sxs-lookup"><span data-stu-id="5f22f-143">When the user clicks the **Get Addresses** button, the `myGetAddresses` event handler obtains an array of addresses from the [addresses](/javascript/api/outlook/office.entities#addresses) property of the `_MyEntities` object, if any address was extracted.</span></span> <span data-ttu-id="5f22f-144">抽出された各住所は、配列内の文字列として保存されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-144">Each extracted address is stored as a string in the array.</span></span> <span data-ttu-id="5f22f-145">`myGetAddresses` はローカル HTML 文字列を `htmlText` で生成し、抽出された住所のリストを表示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-145">`myGetAddresses` forms a local HTML string in `htmlText` to display the list of extracted addresses.</span></span> <span data-ttu-id="5f22f-146">関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-146">The following is the related JavaScript code.</span></span>


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-contact-information"></a><span data-ttu-id="5f22f-147">連絡先情報の抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-147">Extracting contact information</span></span>


<span data-ttu-id="5f22f-p112">ユーザーが **[Get Contact Information]** ボタンをクリックすると、`myGetContacts` イベント ハンドラーが `_MyEntities` オブジェクトの [contacts](/javascript/api/outlook/office.entities#contacts) プロパティから連絡先の配列をそれぞれの情報と共に取得します (連絡先が抽出されていた場合)。抽出された各連絡先は、[Contact](/javascript/api/outlook/office.contact) オブジェクトとして配列に格納されます。`myGetContacts` は、各連絡先に関する詳細なデータを取得します。Outlook がアイテムから連絡先を抽出できるかどうかはコンテキスト次第であることに注意してください。電子メール メッセージの末尾の署名、または少なくとも次のいくつかの情報が連絡先の周辺に存在している必要があります。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p112">When the user clicks the **Get Contact Information** button, the `myGetContacts` event handler obtains an array of contacts together with their information from the [contacts](/javascript/api/outlook/office.entities#contacts) property of the `_MyEntities` object, if any was extracted. Each extracted contact is stored as a [Contact](/javascript/api/outlook/office.contact) object in the array. `myGetContacts` obtains further data about each contact. Note that the context determines whether Outlook can extract a contact from an item&mdash;a signature at the end of an email message, or at least some of the following information would have to exist in the vicinity of the contact:</span></span>


- <span data-ttu-id="5f22f-152">[Contact.personName](/javascript/api/outlook/office.contact#personname) プロパティから取得される連絡先の名前を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-152">The string representing the contact's name from the [Contact.personName](/javascript/api/outlook/office.contact#personname) property.</span></span>

- <span data-ttu-id="5f22f-153">[Contact.businessName](/javascript/api/outlook/office.contact#businessname) プロパティから取得される連絡先に関連付けられた会社名を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-153">The string representing the company name associated with the contact from the [Contact.businessName](/javascript/api/outlook/office.contact#businessname) property.</span></span>

- <span data-ttu-id="5f22f-p113">[Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) プロパティから取得される、連絡先に関連付けられた電話番号の配列。各電話番号は [PhoneNumber](/javascript/api/outlook/office.phonenumber) オブジェクトによって表されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p113">The array of telephone numbers associated with the contact from the [Contact.phoneNumbers](/javascript/api/outlook/office.contact#phonenumbers) property. Each telephone number is represented by a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object.</span></span>

- <span data-ttu-id="5f22f-156">電話番号配列内の **PhoneNumber** メンバーごとの、[PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) プロパティから取得される電話番号を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-156">For each **PhoneNumber** member in the telephone numbers array, the string representing the telephone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="5f22f-p114">[Contact.urls](/javascript/api/outlook/office.contact#urls) プロパティから取得される連絡先に関連付けられた URL の配列。各 URL は配列メンバーの文字列として表されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p114">The array of URLs associated with the contact from the [Contact.urls](/javascript/api/outlook/office.contact#urls) property. Each URL is represented as a string in an array member.</span></span>

- <span data-ttu-id="5f22f-p115">[Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) プロパティから取得される、連絡先に関連付けられた電子メール アドレスの配列。各電子メール アドレスは配列メンバーの文字列として表されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p115">The array of email addresses associated with the contact from the [Contact.emailAddresses](/javascript/api/outlook/office.contact#emailaddresses) property. Each email address is represented as a string in an array member.</span></span>

- <span data-ttu-id="5f22f-p116">[Contact.addresses](/javascript/api/outlook/office.contact#addresses) プロパティから取得される、連絡先に関連付けられた郵送先住所の配列。各郵送先住所は配列メンバーの文字列として表されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p116">The array of postal addresses associated with the contact from the [Contact.addresses](/javascript/api/outlook/office.contact#addresses) property. Each postal address is represented as a string in an array member.</span></span>

<span data-ttu-id="5f22f-p117">`myGetContacts` はローカル HTML 文字列を `htmlText` で生成し、各連絡先のデータを表示します。関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p117">`myGetContacts` forms a local HTML string in `htmlText` to display the data for each contact. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
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
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-email-addresses"></a><span data-ttu-id="5f22f-165">電子メール アドレスの抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-165">Extracting email addresses</span></span>


<span data-ttu-id="5f22f-p118">ユーザーが **[Get Email Addresses]** ボタンをクリックすると、`myGetEmailAddresses` イベント ハンドラーが `_MyEntities` オブジェクトの [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) プロパティから SMTP 電子メール アドレスの配列を取得します (メール アドレスが抽出されていた場合)。抽出された各電子メール アドレスは、文字列として配列に格納されます。`myGetEmailAddresses` はローカル HTML 文字列を `htmlText` で生成し、抽出された電子メール アドレスの一覧を表示します。関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p118">When the user clicks the **Get Email Addresses** button, the `myGetEmailAddresses` event handler obtains an array of SMTP email addresses from the [emailAddresses](/javascript/api/outlook/office.entities#emailaddresses) property of the `_MyEntities` object, if any was extracted. Each extracted email address is stored as a string in the array. `myGetEmailAddresses` forms a local HTML string in `htmlText` to display the list of extracted email addresses. The following is the related JavaScript code.</span></span>


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## <a name="extracting-meeting-suggestions"></a><span data-ttu-id="5f22f-170">会議提案の抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-170">Extracting meeting suggestions</span></span>


<span data-ttu-id="5f22f-171">ユーザーが **[Get Meeting Suggestions]** ボタンをクリックすると、`myGetMeetingSuggestions` イベント ハンドラーが `_MyEntities` オブジェクトの [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) プロパティから会議提案の配列を取得します (会議提案が抽出されていた場合)。</span><span class="sxs-lookup"><span data-stu-id="5f22f-171">When the user clicks the **Get Meeting Suggestions** button, the `myGetMeetingSuggestions` event handler obtains an array of meeting suggestions from the [meetingSuggestions](/javascript/api/outlook/office.entities#meetingsuggestions) property of the `_MyEntities` object, if any was extracted.</span></span>


 > [!NOTE]
 > <span data-ttu-id="5f22f-172">**MeetingSuggestion** エンティティ型をサポートしているのはメッセージだけであり、予定ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5f22f-172">Only messages but not appointments support the **MeetingSuggestion** entity type.</span></span>

<span data-ttu-id="5f22f-p119">抽出された各会議提案は、[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) オブジェクトとして配列に格納されます。`myGetMeetingSuggestions` は、各会議提案に関する次の詳細なデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p119">Each extracted meeting suggestion is stored as a [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object in the array. `myGetMeetingSuggestions` obtains further data about each meeting suggestion:</span></span>


- <span data-ttu-id="5f22f-175">[MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) プロパティから取得される会議提案として識別された文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-175">The string that was identified as a meeting suggestion from the [MeetingSuggestion.meetingString](/javascript/api/outlook/office.meetingsuggestion#meetingstring) property.</span></span>

- <span data-ttu-id="5f22f-p120">[MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) プロパティから取得される、会議の出席者の配列。各出席者は [EmailUser](/javascript/api/outlook/office.emailuser) オブジェクトによって表されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p120">The array of meeting attendees from the [MeetingSuggestion.attendees](/javascript/api/outlook/office.meetingsuggestion#attendees) property. Each attendee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="5f22f-178">出席者ごとの、[EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) プロパティから取得される名前。</span><span class="sxs-lookup"><span data-stu-id="5f22f-178">For each attendee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="5f22f-179">出席者ごとの、[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) プロパティから取得される SMTP アドレス。</span><span class="sxs-lookup"><span data-stu-id="5f22f-179">For each attendee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

- <span data-ttu-id="5f22f-180">[MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) プロパティから取得される、会議提案の場所を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-180">The string representing the location of the meeting suggestion from the [MeetingSuggestion.location](/javascript/api/outlook/office.meetingsuggestion#location) property.</span></span>

- <span data-ttu-id="5f22f-181">[MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) プロパティから取得される、会議提案の議題を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-181">The string representing the subject of the meeting suggestion from the [MeetingSuggestion.subject](/javascript/api/outlook/office.meetingsuggestion#subject) property.</span></span>

- <span data-ttu-id="5f22f-182">[MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) プロパティから取得される、会議提案の開始時刻を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-182">The string representing the start time of the meeting suggestion from the [MeetingSuggestion.start](/javascript/api/outlook/office.meetingsuggestion#start) property.</span></span>

- <span data-ttu-id="5f22f-183">[MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) プロパティから取得される会議提案の終了時刻を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-183">The string representing the end time of the meeting suggestion from the [MeetingSuggestion.end](/javascript/api/outlook/office.meetingsuggestion#end) property.</span></span>

<span data-ttu-id="5f22f-p121">`myGetMeetingSuggestions` はローカル HTML 文字列を `htmlText` で生成し、会議提案ごとのデータを表示します。関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p121">`myGetMeetingSuggestions` forms a local HTML string in `htmlText` to display the data for each of the meeting suggestions. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
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


## <a name="extracting-phone-numbers"></a><span data-ttu-id="5f22f-186">電話番号の抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-186">Extracting phone numbers</span></span>


<span data-ttu-id="5f22f-p122">ユーザーが **[Get Phone Numbers]** ボタンをクリックすると、`myGetPhoneNumbers` イベント ハンドラーが `_MyEntities` オブジェクトの [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) プロパティから電話番号の配列を取得します (電話番号が抽出されていた場合)。抽出された各電話番号は、[PhoneNumber](/javascript/api/outlook/office.phonenumber) オブジェクトとして配列に格納されます。`myGetPhoneNumbers` は、各電話番号に関する次の詳細なデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p122">When the user clicks the **Get Phone Numbers** button, the `myGetPhoneNumbers` event handler obtains an array of phone numbers from the [phoneNumbers](/javascript/api/outlook/office.entities#phonenumbers) property of the `_MyEntities` object, if any was extracted. Each extracted phone number is stored as a [PhoneNumber](/javascript/api/outlook/office.phonenumber) object in the array. `myGetPhoneNumbers` obtains further data about each phone number:</span></span>


- <span data-ttu-id="5f22f-190">[PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) プロパティから取得される電話番号の種類 (自宅の電話番号など) を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-190">The string representing the kind of phone number, for example, home phone number, from the [PhoneNumber.type](/javascript/api/outlook/office.phonenumber#type) property.</span></span>

- <span data-ttu-id="5f22f-191">[PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) プロパティから取得される、実際の電話番号を表す文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-191">The string representing the actual phone number from the [PhoneNumber.phoneString](/javascript/api/outlook/office.phonenumber#phonestring) property.</span></span>

- <span data-ttu-id="5f22f-192">[PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) プロパティから取得される電話番号として最初に識別された文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-192">The string that was originally identified as the phone number from the [PhoneNumber.originalPhoneString](/javascript/api/outlook/office.phonenumber#originalphonestring) property.</span></span>

<span data-ttu-id="5f22f-p123">`myGetPhoneNumbers` はローカル HTML 文字列を `htmlText` で生成し、各電話番号のデータを表示します。関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p123">`myGetPhoneNumbers` forms a local HTML string in `htmlText` to display the data for each of the phone numbers. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
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


## <a name="extracting-task-suggestions"></a><span data-ttu-id="5f22f-195">タスクの提案の抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-195">Extracting task suggestions</span></span>


<span data-ttu-id="5f22f-p124">ユーザーが **[Get Task Suggestions]** ボタンをクリックすると、`myGetTaskSuggestions` イベント ハンドラーが `_MyEntities` オブジェクトの [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) プロパティからタスクの提案の配列を取得します (タスクの提案が抽出されていた場合)。抽出された各タスクの提案は、[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) オブジェクトとして配列に格納されます。`myGetTaskSuggestions` は、各タスクの提案に関する次の詳細なデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p124">When the user clicks the **Get Task Suggestions** button, the `myGetTaskSuggestions` event handler obtains an array of task suggestions from the [taskSuggestions](/javascript/api/outlook/office.entities#tasksuggestions) property of the `_MyEntities` object, if any was extracted. Each extracted task suggestion is stored as a [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object in the array. `myGetTaskSuggestions` obtains further data about each task suggestion:</span></span>


- <span data-ttu-id="5f22f-199">[TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) プロパティから取得されるタスクの提案として最初に識別された文字列。</span><span class="sxs-lookup"><span data-stu-id="5f22f-199">The string that was originally identified a task suggestion from the [TaskSuggestion.taskString](/javascript/api/outlook/office.tasksuggestion#taskstring) property.</span></span>

- <span data-ttu-id="5f22f-p125">[TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) プロパティから取得される、タスクの割り当て先の配列。各割り当て先は [EmailUser](/javascript/api/outlook/office.emailuser) オブジェクトによって表されます。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p125">The array of task assignees from the [TaskSuggestion.assignees](/javascript/api/outlook/office.tasksuggestion#assignees) property. Each assignee is represented by an [EmailUser](/javascript/api/outlook/office.emailuser) object.</span></span>

- <span data-ttu-id="5f22f-202">割り当て先ごとの、[EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) プロパティから取得される名前。</span><span class="sxs-lookup"><span data-stu-id="5f22f-202">For each assignee, the name from the [EmailUser.displayName](/javascript/api/outlook/office.emailuser#displayname) property.</span></span>

- <span data-ttu-id="5f22f-203">割り当て先ごとの、[EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) プロパティから取得される SMTP アドレス。</span><span class="sxs-lookup"><span data-stu-id="5f22f-203">For each assignee, the SMTP address from the [EmailUser.emailAddress](/javascript/api/outlook/office.emailuser#emailaddress) property.</span></span>

<span data-ttu-id="5f22f-p126">`myGetTaskSuggestions` はローカル HTML 文字列を `htmlText` で生成し、タスクの提案ごとのデータを表示します。関連する JavaScript コードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p126">`myGetTaskSuggestions` forms a local HTML string in `htmlText` to display the data for each task suggestion. The following is the related JavaScript code.</span></span>




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
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


## <a name="extracting-urls"></a><span data-ttu-id="5f22f-206">URL の抽出</span><span class="sxs-lookup"><span data-stu-id="5f22f-206">Extracting URLs</span></span>


<span data-ttu-id="5f22f-p127">ユーザーが **[Get URLs]** ボタンをクリックすると、`myGetUrls` イベント ハンドラーが `_MyEntities` オブジェクトの [urls](/javascript/api/outlook/office.entities#urls) プロパティから URL の配列を取得します (URL が抽出されていた場合)。抽出された各 URL は、文字列として配列に格納されます。`myGetUrls` はローカル HTML 文字列を `htmlText` で生成し、抽出された URL の一覧を表示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p127">When the user clicks the **Get URLs** button, the `myGetUrls` event handler obtains an array of URLs from the [urls](/javascript/api/outlook/office.entities#urls) property of the `_MyEntities` object, if any was extracted. Each extracted URL is stored as a string in the array. `myGetUrls` forms a local HTML string in `htmlText` to display the list of extracted URLs.</span></span>


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="clearing-displayed-entity-strings"></a><span data-ttu-id="5f22f-210">表示されたエンティティ文字列の消去</span><span class="sxs-lookup"><span data-stu-id="5f22f-210">Clearing displayed entity strings</span></span>


<span data-ttu-id="5f22f-p128">最後に、エンティティ アドインでは表示された文字列を消去する `myClearEntitiesBox` イベント ハンドラーを指定しています。関連するコードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-p128">Lastly, the entities add-in specifies a  `myClearEntitiesBox` event handler which clears any displayed strings. The following is the related code.</span></span>


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## <a name="javascript-listing"></a><span data-ttu-id="5f22f-213">JavaScript の内容</span><span class="sxs-lookup"><span data-stu-id="5f22f-213">JavaScript listing</span></span>


<span data-ttu-id="5f22f-214">次に、JavaScript の実装の内容全体を示します。</span><span class="sxs-lookup"><span data-stu-id="5f22f-214">The following is the complete listing of the JavaScript implementation.</span></span>


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
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
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
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
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
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
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
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
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
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
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## <a name="see-also"></a><span data-ttu-id="5f22f-215">関連項目</span><span class="sxs-lookup"><span data-stu-id="5f22f-215">See also</span></span>

- [<span data-ttu-id="5f22f-216">閲覧フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="5f22f-216">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="5f22f-217">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="5f22f-217">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- <span data-ttu-id="5f22f-218">[item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) メソッド</span><span class="sxs-lookup"><span data-stu-id="5f22f-218">[item.getEntities method](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span></span>
