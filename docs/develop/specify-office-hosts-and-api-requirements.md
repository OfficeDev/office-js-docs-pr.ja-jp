---
title: Office のホストと API の要件を指定する
description: アドインが期待Office動作するアプリケーションと API 要件を指定する方法について説明します。
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 07f2505dcfb16bf7000dca01a6d600aac9a63fa0
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253355"
---
# <a name="specify-office-applications-and-api-requirements"></a>Office アプリケーションと API 要件を指定する

アドインOfficeは、特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存して、期待通り動作する場合があります。 たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。

- Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。

- アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> Office アドインが現在サポートされている場所の詳細なビューについては[、「Office](../overview/office-add-in-availability.md)クライアント アプリケーションと Office アドインのプラットフォームの可用性」ページを参照してください。

この記事で説明する中心的な概念を次の表に示します。

|**概念**|**説明**|
|:-----|:-----|
|Office アプリケーション、Office クライアント アプリケーション|アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。|
|プラットフォーム|ブラウザーやOfficeなど、アプリケーションが実行される場所iPad。|
|要件セット|関連する API メンバーの名前付きグループ。 アドインは要件セットを使用して、Officeで使用される API メンバーをサポートするかどうかを判断します。 個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。 要件セットのサポートは、アプリケーションOfficeアプリケーションのバージョンによってOfficeされます。 <br >要件セットはマニフェスト ファイルで指定されます。 マニフェストで要件セットを指定する場合は、アドインを実行するために Office アプリケーションが提供する必要がある API サポートの最小レベルを設定します。 Officeで指定された要件セットをサポートしないアプリケーションではアドインを実行できないので、アドインは [マイ アドイン] に<span class="ui">表示されません</span>。これにより、アドインを使用できる場所が制限されます。 コードでは、ランタイム チェックを使用します。 要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。|
|ランタイム チェック|アドインを実行している Officeがアドインで使用される要件セットまたはメソッドをサポートするかどうかを判断するために実行時に実行されるテスト。 ランタイム チェックを実行するには、メソッド、要件セット、または要件セットの一部ではないメソッド名を持つ **if** ステートメント `isSetSupported` を使用します。 ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。 要件セットとは異なり、ランタイム チェックでは、Office アプリケーションがアドインを実行するために提供する必要がある最小レベルの API サポートは指定されません。 代わりに **、if** ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。 サポートされている場合には、アドインで追加機能を提供できます。 ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。|

## <a name="before-you-begin"></a>始める前に

アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。 アドインでランタイム チェックを使用する場合は、最新の JavaScript API (Office) ライブラリをoffice.jsしてください。

### <a name="specify-the-latest-add-in-manifest-schema"></a>最新のアドイン マニフェスト スキーマを指定する

アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。 アドイン マニフェスト [の OfficeApp](../reference/manifest/officeapp.md) 要素を次のように設定します。 次の使用例は、型を示 `TaskPaneApp` しています。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>JavaScript API ライブラリOfficeを指定する

ランタイム チェックを使用する場合は、コンテンツ配信ネットワーク (Office) から JavaScript API ライブラリの最新バージョンを参照CDN。 その場合、HTML に次の `script` タグを追加します。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a>アプリケーションまたは API Officeを指定するオプション

アプリケーションまたは API Officeを指定する場合、考慮すべきいくつかの要因があります。 次の図に、アドインで使用すべき手法の判別方法を示します。

![アプリケーションまたは API の要件を指定するときに、アドインに最適なOfficeを選択する](../images/options-for-office-hosts.png)

- アドインが 1 つのアプリケーションでOffice場合は、マニフェスト `Hosts` で要素を設定します。 詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。

- アドインを実行するためにOfficeアプリケーションでサポートする必要がある最小要件セットまたは API メンバーを設定するには、マニフェストで要素 `Requirements` を設定します。 詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。

- 特定の要件セットまたは API メンバーが Office アプリケーションで使用できる場合は、追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。 たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。 詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。

## <a name="set-the-hosts-element"></a>Hosts 要素を設定する

アドインを 1 つのクライアント アプリケーションでOfficeするには、マニフェストの `Hosts` and `Host` 要素を使用します。 要素を指定しない場合、アドインは、指定した種類 (メール、作業ウィンドウ、またはコンテンツ) でサポートされているすべての Office アプリケーションで `Hosts` `OfficeApp` 実行されます。

たとえば、次の宣言と宣言は、アドインが Excel のすべてのリリース (Excel on the web、Windows、および iPad を含む) で動作 `Hosts` `Host` iPad。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

要素 `Hosts` には、1 つ以上の要素を含 `Host` めできます。 要素 `Host` は、アドインOffice必要なアプリケーションを指定します。 属性 `Name` は必須であり、次のいずれかの値に設定できます。

| 名前          | Office クライアント アプリケーション                     | 使用可能なアドインの種類 |
|:--------------|:-----------------------------------------------|:-----------------------|
| データベース      | Access Web アプリ                                | 作業ウィンドウ              |
| Document      | Word on the web、Windows、Mac、iPad            | 作業ウィンドウ              |
| Mailbox       | Outlook、Windows、Mac、Android、iOS | メール                   |
| Notebook      | OneNote on the web                             | 作業ウィンドウ、コンテンツ     |
| Presentation  | PowerPoint on the web、Windows、Mac、iPad      | 作業ウィンドウ、コンテンツ     |
| Project       | Windows での Project                             | 作業ウィンドウ              |
| Workbook      | Excel on the web、Windows、Mac、iPad           | 作業ウィンドウ、コンテンツ     |

> [!NOTE]
> この `Name` 属性は、アドインOffice実行できるクライアント アプリケーションの名前を指定します。 Officeアプリケーションは、さまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、およびモバイル デバイスで実行されます。 アドインを実行するために使用するプラットフォームを指定することはできません。 たとえば、指定した場合は、web OutlookとWindowsの両方をアドインの実行 `Mailbox` に使用できます。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="set-the-requirements-element-in-the-manifest"></a>マニフェストで Requirements 要素を設定する

要素は、アドインを実行するために、Officeアプリケーションでサポートする必要がある最小要件セットまたは API メンバー `Requirements` を指定します。 要素 `Requirements` は、アドインで使用される要件セットと個々のメソッドの両方を指定できます。 アドイン マニフェスト スキーマのバージョン 1.1 では、アドインを除くすべてのアドインの要素 `Requirements` Outlookです。

> [!WARNING]
> 要素を使用 `Requirements` して、アドインで使用する必要がある重要な要件セットまたは API メンバーのみを指定します。 Office アプリケーションまたはプラットフォームが要素で指定された要件セットまたは API メンバーをサポートしない場合、アドインは、そのアプリケーションまたはプラットフォームでは実行されません。また、My アドインには `Requirements` **表示** されません。代わりに、Office アプリケーションのすべてのプラットフォーム (Excel on the web、Windows、iPad など) でアドインを使用iPad。 すべてのアプリケーションとプラットフォームでアドインをOfficeするには、要素の代わりにランタイム チェックを使用 `Requirements` します。

次のコード例は、次をサポートしているすべてのクライアント アプリケーションでOfficeアドインを示しています。

-  `TableBindings` 要件セット 。最小バージョンは "1.1" です。

-  `OOXML` 要件セット 。最小バージョンは "1.1" です。

-  `Document.getSelectedDataAsync` メソッド。

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

- 要素 `Requirements` には、子要素 `Sets` と子 `Methods` 要素が含まれます。

- 要素 `Sets` には、1 つ以上の要素を含 `Set` めできます。 `DefaultMinVersion` すべての子要素の `MinVersion` 既定値を指定 `Set` します。

- 要素 `Set` は、アドインを実行するためにOfficeアプリケーションがサポートする必要がある要件セットを指定します。 属性 `Name` は、要件セットの名前を指定します。 要件 `MinVersion` セットの最小バージョンを指定します。 `MinVersion`overrides の値 API メンバーが属する要件セットと要件セットのバージョンの詳細については、「Officeアドイン要件セット」 `DefaultMinVersion` [を参照してください](../reference/requirement-sets/office-add-in-requirement-sets.md)。

- 要素 `Methods` には、1 つ以上の要素を含 `Method` めできます。 アドインで要素を `Methods` 使用Outlookすることはできません。

- 要素は、アドインが実行されるアプリケーションでサポートされる必要Office個別 `Method` のメソッドを指定します。 属性 `Name` は必須であり、親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="use-runtime-checks-in-your-javascript-code"></a>JavaScript コードでランタイム チェックを使用する

特定の要件セットがアプリケーションでサポートされている場合は、アドインに追加の機能を提供Officeがあります。 たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。 その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。 `isSetSupported`実行時に、アドインを実行Officeアプリケーションが要件セットをサポートするかどうかを判断します。 要件セットがサポートされている場合は、true を返し、その要件セットの API メンバーを使用する追加のコード `isSetSupported` を実行します。  アプリケーションがOfficeが要件セットをサポートしない場合 `isSetSupported` **、false** を返し、追加のコードは実行されません。 次のコードは `isSetSupported` と共に使用する構文を示しています。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。 利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。
- _MinimumVersion_ (省略可能) は、ステートメント内のコードを実行するために Office アプリケーションがサポートする必要がある最小要件セット のバージョンを指定する文字列 `if` です (たとえば **、"1.9")。**

> [!WARNING]
> メソッドを呼び `isSetSupported` 出す場合、パラメーターの値 (指定されている場合 `MinimumVersion` ) は文字列である必要があります。 これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。
> `number` のオーバーロードは非推奨になります。

次 `isSetSupported` のように、 `RequirementSetName` アプリケーションに関連付Office使用します。

|Office アプリケーション|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

これらの `isSetSupported` アプリケーションのメソッドと要件セットは、アプリケーションの最新のOffice.jsで使用CDN。 アドインから例外をOffice.js場合CDN、未定義のため、アドインで `isSetSupported` 例外が生成される場合があります。 詳細については[、「JavaScript API ライブラリの最新のOffice指定する」を参照してください](#specify-the-latest-office-javascript-api-library)。

次のコード例は、アドインが異なる要件セットまたは API メンバーをサポートする可能性Officeアプリケーションに対して異なる機能を提供する方法を示しています。

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

API の一部のメンバーは、要件のセットに属していません。 これは[、Office JavaScript API](../reference/javascript-api-for-office.md)名前空間の一部である API メンバー (Outlook メールボックス API を除くすべての API) にのみ適用されますが `Office.` [、Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (内の何でも) [](/javascript/api/outlook) `Word.` [、Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (内の何でも)、または OneNote `Excel.` [JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) ( `OneNote.` 何でも) 名前空間に属する API メンバーには適用されません。 アドインが要件セットの一部ではないメソッドに依存している場合は、ランタイム チェックを使用して、次のコード例に示すように、メソッドが Office アプリケーションでサポートされているかどうかを判断できます。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例は、アプリケーションがサポートOfficeチェックします `document.setSelectedDataAsync` 。

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
