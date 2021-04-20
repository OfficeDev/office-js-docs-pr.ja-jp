---
title: Office のホストと API の要件を指定する
description: アドインが期待Office動作するために必要なアプリケーションと API の要件を指定する方法について説明します。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 948e86e99150ebf2d0bc7deaa5512627679b025f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237841"
---
# <a name="specify-office-applications-and-api-requirements"></a>Office アプリケーションと API 要件を指定する

アドインOffice期待通り動作するために、特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存している場合があります。 たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。

- Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。

- アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> Office アドインが現在サポートされている場所の大きなレベルについては、「Office クライアント アプリケーションと Office[](../overview/office-add-in-availability.md)アドインのプラットフォームの可用性」ページを参照してください。

この記事で説明する中心的な概念を次の表に示します。

|**概念**|**説明**|
|:-----|:-----|
|Office、Office クライアント アプリケーション|アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。|
|プラットフォーム|ブラウザー Office iPad など、アプリケーションが実行される場所。|
|要件セット|関連する API メンバーの名前付きグループ。 アドインは要件セットを使用して、アドインが使用Office API メンバーをサポートするかどうかを判断します。 個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。 要件セットのサポートは、OfficeアプリケーションとアプリケーションのバージョンによってOfficeされます。 <br >要件セットはマニフェスト ファイルで指定されます。 マニフェストで要件セットを指定する場合は、アドインを実行するために Office アプリケーションが提供する必要がある API サポートの最小レベルを設定します。 Officeで指定された要件セットをサポートしないアプリケーションはアドインを実行できないので、アドインは [マイ アドイン] に <span class="ui">表示されません</span>。これにより、アドインを使用できる場所が制限されます。 コードでは、ランタイム チェックを使用します。 要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。|
|ランタイム チェック|アドインを実行している Office アプリケーションが、アドインで使用される要件セットまたはメソッドをサポートするかどうかを判断するために実行時に実行されるテスト。 ランタイム チェックを実行するには **、if** ステートメントをメソッド、要件セット、または要件セットの一部ではないメソッド名と一緒 `isSetSupported` に使用します。 ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。 要件セットとは異なり、ランタイム チェックでは、アドインを実行するために Office アプリケーションが提供する必要がある最小レベルの API サポートは指定されません。 代わりに **、if** ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。 サポートされている場合には、アドインで追加機能を提供できます。 ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。|

## <a name="before-you-begin"></a>始める前に

アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。 アドインでランタイム チェックを使用する場合は、最新の JavaScript API (Office) ライブラリを使用office.jsしてください。

### <a name="specify-the-latest-add-in-manifest-schema"></a>最新のアドイン マニフェスト スキーマを指定する

アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。 アドイン マニフェスト `OfficeApp` の要素を次のように設定します。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>JavaScript API ライブラリOffice指定する

ランタイム チェックを使用する場合は、コンテンツ配信ネットワーク (CDN) から Office JavaScript API ライブラリの最新バージョンを参照します。 その場合、HTML に次の `script` タグを追加します。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>アプリケーションまたは API Officeを指定するオプション

アプリケーションまたは API Officeを指定する場合は、いくつかの要素を考慮する必要があります。 次の図に、アドインで使用すべき手法の判別方法を示します。

![アプリケーションまたは API の要件を指定するときに、アドインにOfficeオプションを選択する](../images/options-for-office-hosts.png)

- アドインが 1 つのアプリケーションで実行Officeマニフェスト `Hosts` で要素を設定します。 詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。

- アドインを実行するために Office アプリケーションがサポートする必要がある最小要件セットまたは API メンバーを設定するには、マニフェストで要素 `Requirements` を設定します。 詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。

- 特定の要件セットまたは API メンバーが Office アプリケーションで使用可能な場合に追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。 たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。 詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。

## <a name="set-the-hosts-element"></a>Hosts 要素を設定する

アドインを 1 つのクライアント アプリケーションでOfficeするには、マニフェストの `Hosts` 要素 `Host` を使用します。 要素を指定しない場合、アドインは、アドインによってサポートOfficeのすべての Office アプリケーション `Hosts` で実行されます。

たとえば、次と宣言は、アドインが Excel のリリース `Hosts` (Excel on the web、Windows、iPad など) で動作する場合に指定 `Host` します。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

要素 `Hosts` には、1 つ以上の要素を含 `Host` めできます。 この `Host` 要素は、アドインOffice必要なアプリケーションを指定します。 この `Name` 属性は必須であり、次のいずれかの値に設定できます。

| 名前          | Office クライアント アプリケーション                      |
|:--------------|:----------------------------------------------|
| データベース      | Access Web アプリ                               |
| Document      | Word on the web、Windows、Mac、iPad           |
| メールボックス       | Outlook on the web、Windows、Mac、Android、iOS|
| Presentation  | PowerPoint on the web、Windows、Mac、iPad     |
| Project       | Windows での Project                            |
| Workbook      | Excel on the web、Windows、Mac、iPad          |

> [!NOTE]
> この `Name` 属性は、アドインOffice実行できるクライアント アプリケーションの名前を指定します。 Officeアプリケーションは、さまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行されます。 アドインを実行するために使用するプラットフォームを指定することはできません。 たとえば、指定した場合、Outlook on the web と Windows の両方を使用してアドイン `Mailbox` を実行できます。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="set-the-requirements-element-in-the-manifest"></a>マニフェストで Requirements 要素を設定する

この要素は、アドインを実行するために、Office アプリケーションでサポートする必要がある最小要件セットまたは API メンバー `Requirements` を指定します。 この `Requirements` 要素は、要件セットと、アドインで使用される個々のメソッドの両方を指定できます。 アドイン マニフェスト スキーマのバージョン 1.1 では、Outlook アドインを除くすべてのアドインでこの要素は `Requirements` オプションです。

> [!WARNING]
> この要素は `Requirements` 、アドインで使用する必要がある重要な要件セットまたは API メンバーを指定する場合にのみ使用します。 Office アプリケーションまたはプラットフォームが要素で指定された要件セットまたは API メンバーをサポートしない場合、アドインは、そのアプリケーションまたはプラットフォームでは実行されません。また、マイ アドインには表示されません `Requirements` 。 代わりに、Excel on the web、Windows、iPad など、Office アプリケーションのすべてのプラットフォームでアドインを使用することをお勧めします。 アドインをすべてのアプリケーションとプラットフォーム  _でOfficeするには_ 、要素の代わりにランタイム チェックを使用 `Requirements` します。

次のコード例は、以下をサポートしているすべてのクライアント Office読み込まれるアドインを示しています。

-  `TableBindings` 要件セット。最小バージョンは "1.1" です。

-  `OOXML` 要件セット。最小バージョンは "1.1" です。

-  `Document.getSelectedDataAsync` メソッドを使用します。

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

- 要素 `Requirements` には、子要素 `Sets` `Methods` と子要素が含まれます。

- 要素 `Sets` には、1 つ以上の要素を含 `Set` めできます。 `DefaultMinVersion` すべての子要素の `MinVersion` 既定値を指定 `Set` します。

- この `Set` 要素は、アドインを実行するためにOfficeアプリケーションがサポートする必要がある要件セットを指定します。 この `Name` 属性は、要件セットの名前を指定します。 要件 `MinVersion` セットの最小バージョンを指定します。 `MinVersion` overrides the value of `DefaultMinVersion` Requirement sets and requirement set versions that your API members belong to, see Office [Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).

- 要素 `Methods` には、1 つ以上の要素を含 `Method` めできます。この要素を Outlook `Methods` アドインと一緒に使用することはできません。

- この `Method` 要素は、アドインを実行するアプリケーションでサポートOfficeする必要がある個別のメソッドを指定します。この `Name` 属性は必須であり、親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="use-runtime-checks-in-your-javascript-code"></a>JavaScript コードでランタイム チェックを使用する

特定の要件セットがアプリケーションでサポートされている場合は、アドインに追加の機能をOfficeがあります。 たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。 その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。 `isSetSupported` 実行時に、アドインを実行Officeアプリケーションが要件セットをサポートするかどうかを決定します。 要件セットがサポートされている場合は true を返し、その要件セットの API メンバーを使用する追加コード `isSetSupported` を実行します。  アプリケーションがOfficeが要件セットをサポートしない場合は false を返し、追加のコード `isSetSupported` は実行されません。  次のコードは `isSetSupported` と共に使用する構文を示しています。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。 利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。
- _MinimumVersion_ (オプション) は、ステートメント内のコードを実行するために Office アプリケーションがサポートする必要がある最小要件セットのバージョンを指定する文字列です (例: `if` "**1.9**")。

> [!WARNING]
> メソッドを呼 `isSetSupported` び出す場合、パラメーターの値 (指定されている場合 `MinimumVersion` ) は文字列である必要があります。 これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。
> `number` のオーバーロードは非推奨になります。

次 `isSetSupported` のように、 `RequirementSetName` アプリケーションに関連付Office使用します。

|Office アプリケーション|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

これらの `isSetSupported` アプリケーションのメソッドと要件セットは、CDN の最新の Office.js ファイルで使用できます。 CDN からのカスタム Office.js使用しない場合は、未定義のため、アドインで例外 `isSetSupported` が生成される可能性があります。 詳細については [、「JavaScript API ライブラリの最新のOffice指定する」を参照してください](#specify-the-latest-office-javascript-api-library)。

次のコード例は、さまざまな要件セットまたは API メンバーをサポートする可能性があるさまざまな Office アプリケーションに対して、アドインがさまざまな機能を提供する方法を示しています。

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

API の一部のメンバーは、要件のセットに属していません。 [これは、Office JavaScript API](../reference/javascript-api-for-office.md)名前空間の一部である API メンバー (Outlook メールボックス API を除くすべてのメンバー) にのみ適用されますが、Word JavaScript API に属する API メンバー (次の場合は何でも含む)、Excel JavaScript API (何でも含む `Office.` [](/javascript/api/outlook)[](../reference/overview/word-add-ins-reference-overview.md) `Word.` [](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` )、OneNote [JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (すべての名前空間 `OneNote.` ) には適用されません。 アドインが要件セットの一部ではないメソッドに依存している場合は、ランタイム チェックを使用して、次のコード例に示すように、メソッドが Office アプリケーションでサポートされているかどうかを判断できます。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例では、アプリケーションがサポートOfficeチェックします `document.setSelectedDataAsync` 。

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
