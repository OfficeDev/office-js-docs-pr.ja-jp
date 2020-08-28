---
title: Office のホストと API の要件を指定する
description: アドインが意図したとおりに動作するように Office アプリケーションと API の要件を指定する方法について説明します。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 90ee7c3a5ad01252336608c02f995bbcbbe94212
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292630"
---
# <a name="specify-office-applications-and-api-requirements"></a>Office アプリケーションと API の要件を指定する

Office アドインは、想定どおりに動作するために、特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存している可能性があります。 たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。

- Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。

- アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> 現在、Office アドインが現在サポートされている場所の概要については、「office [アドインの office クライアントアプリケーションとプラットフォームの可用性](../overview/office-add-in-availability.md) 」ページを参照してください。

この記事で説明する中心的な概念を次の表に示します。

|**概念**|**説明**|
|:-----|:-----|
|Office アプリケーション、Office クライアントアプリケーション|アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。|
|プラットフォーム|Office アプリケーションが実行されている場所 (ブラウザーや iPad など)。|
|要件セット|関連する API メンバーの名前付きグループ。 アドインは、要件セットを使用して、Office アプリケーションがアドインで使用される API メンバーをサポートするかどうかを判断します。 個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。 要件セットのサポートは、Office アプリケーションと Office アプリケーションのバージョンによって異なります。 <br >要件セットはマニフェスト ファイルで指定されます。 マニフェストで要件セットを指定するときは、アドインを実行するために Office アプリケーションが提供する必要がある API サポートの最小レベルを設定します。 マニフェストで指定されている要件セットをサポートしていない Office アプリケーションは、アドインを実行できず、アドインは <span class="ui">自分</span>のアドインに表示されません。これにより、アドインを使用できる場所が制限されます。 コードでは、ランタイム チェックを使用します。 要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。|
|ランタイム チェック|実行時に実行されるテストで、アドインを実行している Office アプリケーションが、アドインで使用される要件セットまたはメソッドをサポートしているかどうかを判断します。 ランタイムチェックを実行するには、メソッドの **if** ステートメント `isSetSupported` 、要件セット、または要件セットの一部ではないメソッド名を使用します。 ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。 ランタイムチェックは要件セットとは異なり、Office アプリケーションがアドインを実行するために提供する必要がある API サポートの最小レベルを指定しません。 代わりに、 **if** ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。 サポートされている場合には、アドインで追加機能を提供できます。 ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。|

## <a name="before-you-begin"></a>始める前に

アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。 アドインでランタイムチェックを使用する場合は、最新の Office JavaScript API (office.js) ライブラリを使用していることを確認してください。

### <a name="specify-the-latest-add-in-manifest-schema"></a>最新のアドイン マニフェスト スキーマを指定する

アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。 `OfficeApp`アドインマニフェストの要素を次のように設定します。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>最新の Office JavaScript API ライブラリを指定する

ランタイムチェックを使用する場合は、コンテンツ配信ネットワーク (CDN) から、最新バージョンの Office JavaScript API ライブラリを参照します。 その場合、HTML に次の `script` タグを追加します。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>Office アプリケーションまたは API の要件を指定するオプション

Office アプリケーションまたは API の要件を指定する場合は、考慮すべきいくつかの要因があります。 次の図に、アドインで使用すべき手法の判別方法を示します。

![Office アプリケーションまたは API の要件を指定するときに、アドインに最適なオプションを選択する](../images/options-for-office-hosts.png)

- アドインを1つの Office アプリケーションで実行する場合は、 `Hosts` マニフェスト内の要素を設定します。 詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。

- Office アプリケーションがアドインを実行するためにサポートする必要のある最小要件セットまたは API メンバーを設定するには、 `Requirements` マニフェストで要素を設定します。 詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。

- Office アプリケーションで特定の要件セットや API メンバーが使用可能な場合に追加機能を提供するには、アドインの JavaScript コードでランタイムチェックを実行します。 たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。 詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。

## <a name="set-the-hosts-element"></a>Hosts 要素を設定する

1つの Office クライアントアプリケーションでアドインを実行するには、 `Hosts` マニフェスト内の要素と要素を使用し `Host` ます。 要素を指定しない場合 `Hosts` 、アドインは Office アドインでサポートされているすべての office アプリケーションで実行されます。

たとえば、次の `Hosts` と宣言は、 `Host` アドインが excel のすべてのリリースで動作することを指定します。これには、Web、Windows、iPad 上の excel が含まれます。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

要素には、 `Hosts` 1 つ以上の要素を含めることができ `Host` ます。 要素は、 `Host` アドインが必要とする Office アプリケーションを指定します。 `Name`属性は必須で、次のいずれかの値に設定できます。

| 名前          | Office クライアントアプリケーション                      |
|:--------------|:----------------------------------------------|
| データベース      | Access Web アプリ                               |
| ドキュメント      | Web、Windows、Mac、iPad の Word           |
| メールボックス       | Outlook on the web、Windows、Mac、Android、iOS|
| Presentation  | PowerPoint on the web、Windows、Mac、iPad     |
| Project       | Windows での Project                            |
| Workbook      | Web、Windows、Mac、iPad の Excel          |

> [!NOTE]
> この `Name` 属性は、アドインを実行できる Office クライアントアプリケーションを指定します。 Office アプリケーションは、さまざまなプラットフォームでサポートされており、デスクトップ、web ブラウザー、タブレット、モバイルデバイスで動作します。 アドインを実行するために使用するプラットフォームを指定することはできません。 たとえば、を指定すると `Mailbox` 、web 上の Outlook と Windows の両方を使用してアドインを実行できます。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="set-the-requirements-element-in-the-manifest"></a>マニフェストで Requirements 要素を設定する

要素は、 `Requirements` アドインを実行するために Office アプリケーションでサポートされている必要のある最小要件セットまたは API メンバーを指定します。 要素は、 `Requirements` 要件セットと、アドインで使用される個々のメソッドの両方を指定できます。 アドインマニフェストスキーマのバージョン1.1 では、Outlook アドインを除くすべてのアドインで、この `Requirements` 要素は省略可能です。

> [!WARNING]
> 要素のみを使用して、 `Requirements` アドインで使用する必要がある重要な要件セットまたは API メンバーを指定します。 Office アプリケーションまたはプラットフォームが、要素で指定されている要件セットや API メンバーをサポートしていない場合、 `Requirements` アドインはそのアプリケーションまたはプラットフォームでは実行されず、 **アドイン**には表示されません。その代わりに、Office アプリケーションのすべてのプラットフォーム (web、Windows、iPad など) でアドインを使用できるようにすることをお勧めします。 _すべて_の Office アプリケーションおよびプラットフォームでアドインを使用できるようにするには、要素の代わりにランタイムチェックを使用し `Requirements` ます。

次のコード例は、次のものをサポートするすべての Office クライアントアプリケーションで読み込まれるアドインを示しています。

-  `TableBindings` 要件セット。最小バージョンは "1.1" です。

-  `OOXML` 要件セット。最小バージョンは "1.1" です。

-  `Document.getSelectedDataAsync` 手段.

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- `Requirements`要素にはおよび子要素が含まれてい `Sets` `Methods` ます。

- 要素には、 `Sets` 1 つ以上の要素を含めることができ `Set` ます。 `DefaultMinVersion``MinVersion`すべての子要素の既定値を指定し `Set` ます。

- 要素は、 `Set` Office アプリケーションがアドインを実行するためにサポートする必要がある要件セットを指定します。 属性は、 `Name` 要件セットの名前を指定します。 は、 `MinVersion` 要件セットの最小バージョンを指定します。 `MinVersion``DefaultMinVersion`API メンバーが属する要件セットと要件セットのバージョンの詳細については、「 [Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。

- 要素には、 `Methods` 1 つ以上の要素を含めることができ `Method` ます。Outlook アドインで要素を使用することはできません `Methods` 。

- 要素は、 `Method` アドインを実行する Office アプリケーションでサポートする必要がある個別のメソッドを指定します。この `Name` 属性は必須で、その親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="use-runtime-checks-in-your-javascript-code"></a>JavaScript コードでランタイム チェックを使用する

Office アプリケーションで特定の要件セットがサポートされている場合は、アドインに追加機能を提供することをお勧めします。 たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。 その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。 `isSetSupported` 実行時に、アドインを実行している Office アプリケーションが要件セットをサポートするかどうかを指定します。 要件セットがサポートされている場合は、 `isSetSupported` **true** を返し、その要件セットの API メンバーを使用する追加のコードを実行します。 Office アプリケーションが要件セットをサポートしていない場合は、 `isSetSupported` **false** を返し、追加のコードは実行されません。 次のコードは `isSetSupported` と共に使用する構文を示しています。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。 利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。
- _MinimumVersion_ (省略可能) は、ステートメント内でコードを実行するために Office アプリケーションがサポートする必要がある最小要件セットのバージョンを指定する文字列です `if` ("**1.9**" など)。

> [!WARNING]
> メソッドを呼び出す場合 `isSetSupported` 、 `MinimumVersion` パラメーター (指定されている場合) の値は文字列である必要があります。 これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。
> `number` のオーバーロードは非推奨になります。

`isSetSupported` `RequirementSetName` Office アプリケーションに関連付けられているを次のように使用します。

|Office アプリケーション|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

`isSetSupported`これらのアプリケーションのためのメソッドと要件セットは、CDN の最新の Office.js ファイルで入手できます。 CDN から Office.js を使用しない場合、アドインは未定義となるため、例外が生成されることがあり `isSetSupported` ます。 詳細については、「 [最新の Office JAVASCRIPT API ライブラリを指定する](#specify-the-latest-office-javascript-api-library)」を参照してください。

次のコード例は、アドインがさまざまな要件セットや API メンバーをサポートする可能性のある、さまざまな Office アプリケーションに対してさまざまな機能を提供する方法を示しています。

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
    // Run code that provides additional functionality using the Word JavaScript API when the add-in runs in Word 2016 or later.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office application is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>要件セットにないメソッドを使用したランタイム チェック

API の一部のメンバーは、要件のセットに属していません。 これは、 [Office JAVASCRIPT api](../reference/javascript-api-for-office.md)名前空間 ( `Office.` [Outlook メールボックス api](/javascript/api/outlook)以外のすべて) に属する api メンバーではなく、 [Word javascript api](../reference/overview/word-add-ins-reference-overview.md) (すべてのもの)、Excel javascript api (すべての場合) `Word.` [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` 、または[OneNote javascript api](../reference/overview/onenote-add-ins-javascript-reference.md) (あらゆる場合) の名前空間に含まれる api メンバーにのみ適用され `OneNote.` ます。 アドインが要件セットの一部ではないメソッドに依存している場合は、次のコード例に示すように、ランタイムチェックを使用して、メソッドが Office アプリケーションでサポートされているかどうかを判断できます。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例では、Office アプリケーションがサポートしているかどうかを確認し `document.setSelectedDataAsync` ます。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](add-in-manifests.md)
- [Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
