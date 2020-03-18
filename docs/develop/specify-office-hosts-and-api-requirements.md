---
title: Office のホストと API の要件を指定する
description: アドインが意図したとおりに動作するように Office のホストと API の要件を指定する方法について説明します。
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: ab9b97f3d3232339010179097e1fd03dbeb86aa2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718812"
---
# <a name="specify-office-hosts-and-api-requirements"></a>Office のホストと API の要件を指定する

期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。

- Office の一部のバージョンでのみ利用できる JavaScript API を使用します。たとえば、Excel 2016 で実行するアドインでは、Excel JavaScript API を使用することがあります。

- アドインが使用する API メンバーをサポートするバージョンの Office でのみ実行します。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> 現時点での Office アドインのサポート状況の概要については、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」のページを参照してください。

この記事で説明する中心的な概念を次の表に示します。

|**概念**|**説明**|
|:-----|:-----|
|Office アプリケーション、Office ホスト アプリケーション、Office ホスト、またはホスト|アドインの実行に使用される Office アプリケーション。たとえば、Word や Excel など。|
|プラットフォーム|Office ホストを実行する場所。ブラウザーや iPad など。|
|要件セット|関連する API メンバーの名前付きグループ。アドインは要件セットを使用して、Office ホストが、アドインによって使用される API メンバーをサポートしているかどうかを判別します。個々の API メンバーのサポートをテストするよりも、要件セットのサポートをテストするほうが簡単です。要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります。 <br >要件セットはマニフェスト ファイルで指定されます。 マニフェストで要件セットを指定するときは、アドインを実行するために Office ホストが提供する必要のある最小レベルの API サポートを設定します。 マニフェストで指定されている要件セットをサポートしていない Office ホストはアドインを実行できず、アドインは <span class="ui">[個人用アドイン]</span> に表示されません。これにより、アドインが利用できる場所が制限されます。 コードでは、ランタイム チェックを使用します。 要件セットの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。|
|ランタイム チェック|アドインを実行している Office ホストが、アドインで使用されている要件セットまたはメソッドをサポートしているかどうかを判別するために実行時に行われるテスト。 ランタイムチェックを実行するには、 `isSetSupported`メソッドの**if**ステートメント、要件セット、または要件セットの一部ではないメソッド名を使用します。 ランタイム チェックを使用すると、アドインを、最も多くのお客様が利用できるものにできます。 要件セットとは異なり、ランタイム チェックでは、対象アドインを実行するために Office ホストが提供する必要のある最小レベルの API サポートは指定しません。 代わりに、 **if**ステートメントを使用して、API メンバーがサポートされているかどうかを判断します。 サポートされている場合には、アドインで追加機能を提供できます。 ランタイム チェックを使用するときは、自分のアドインは必ず **[個人用アドイン]** に表示されます。|

## <a name="before-you-begin"></a>始める前に

アドインで最新バージョンのアドイン マニフェスト スキーマを使用する必要があります。 アドインでランタイムチェックを使用する場合は、最新の Office JavaScript API (office .js) ライブラリを使用していることを確認してください。

### <a name="specify-the-latest-add-in-manifest-schema"></a>最新のアドイン マニフェスト スキーマを指定する

アドインのマニフェストでは、アドイン マニフェスト スキーマのバージョン 1.1 を使用する必要があります。 アドインマニフェスト`OfficeApp`の要素を次のように設定します。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a>最新の Office JavaScript API ライブラリを指定する

ランタイムチェックを使用する場合は、コンテンツ配信ネットワーク (CDN) から、最新バージョンの Office JavaScript API ライブラリを参照します。 その場合、HTML に次の `script` タグを追加します。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>Office のホストや API の要件を指定するオプション

Office ホストまたは API の要件を指定するときに、検討すべき事項がいくつかあります。次の図に、アドインで使用すべき手法の判別方法を示します。

![Office のホストまたは API の要件を指定する際に、アドインに最適なオプションを選択する](../images/options-for-office-hosts.png)

- アドインが1つの Office ホストで実行される`Hosts`場合は、マニフェスト内の要素を設定します。 詳しくは、「[Hosts 要素を設定する](#set-the-hosts-element)」を参照してください。

- Office ホストがアドインを実行するためにサポートする必要のある最小要件セットまたは API メンバーを設定する`Requirements`には、マニフェストで要素を設定します。 詳しくは、「[マニフェストで Requirements 要素を設定する](#set-the-requirements-element-in-the-manifest)」をご覧ください。

- Office ホストで特定の要件セットまたは API メンバーが利用可能である場合に追加の機能を提供する場合は、アドインの JavaScript コードでランタイム チェックを実行します。 たとえば、アドインが Excel 2016 で機能する場合は、Excel JavaScript API の API メンバーを使用して追加の機能を提供します。 詳細については、「[JavaScript コードでランタイム チェックを使用する](#use-runtime-checks-in-your-javascript-code)」をご覧ください。

## <a name="set-the-hosts-element"></a>Hosts 要素を設定する

1つの Office ホストアプリケーションでアドインを実行するには、マニフェスト`Hosts`内`Host`の要素と要素を使用します。 `Hosts`要素を指定しない場合、アドインはすべてのホストで実行されます。

たとえば、次`Hosts`のと`Host`宣言は、アドインが excel のすべてのリリースで動作することを指定します。これには、Web、Windows、iPad 上の excel が含まれます。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

要素`Hosts`には、1つ以上`Host`の要素を含めることができます。 要素`Host`は、アドインが必要とする Office ホストを指定します。 `Name`属性は必須で、次のいずれかの値に設定できます。

| 名前          | Office ホスト アプリケーション                                                                  |
|:--------------|:------------------------------------------------------------------------------------------|
| データベース      | Access Web アプリ                                                                           |
| ドキュメント      | Windows 用 Word、Mac 用 Word、iPad 用 Word、Word on the web                               |
| Mailbox       | Outlook on Windows、Outlook on Mac、Outlook on the web、Outlook on Android、Outlook on iOS|
| Presentation  | Windows 用 PowerPoint、Mac 用 PowerPoint、iPad 用 PowerPoint、PowerPoint on the web       |
| Project       | Windows での Project                                                                        |
| Workbook      | Windows 用 Excel、Mac 用 Excel、iPad 用 Excel、Excel on the web                           |

> [!NOTE]
> この`Name`属性は、アドインを実行できる Office ホストアプリケーションを指定します。 Office ホストはさまざまなプラットフォームに対応しており、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行できます。 アドインを実行するために使用するプラットフォームを指定することはできません。 たとえば、`Mailbox` を指定した場合は、アドインの実行に Windows 用 Outlook と Outlook on the web の両方を使用できます。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。


## <a name="set-the-requirements-element-in-the-manifest"></a>マニフェストで Requirements 要素を設定する

要素`Requirements`は、アドインを実行するために Office ホストでサポートされている必要のある最小要件セットまたは API メンバーを指定します。 要素`Requirements`は、要件セットと、アドインで使用される個々のメソッドの両方を指定できます。 アドインマニフェストスキーマのバージョン1.1 では、Outlook アドインを`Requirements`除くすべてのアドインで、この要素は省略可能です。

> [!WARNING]
> `Requirements`要素のみを使用して、アドインで使用する必要がある重要な要件セットまたは API メンバーを指定します。 Office ホストまたはプラットフォームが、 `Requirements`要素で指定されている要件セットや API メンバーをサポートしていない場合、アドインはそのホストまたはプラットフォームでは実行されず、**アドイン**には表示されません。代わりに、web、Windows、iPad の Excel など、Office ホストのすべてのプラットフォームでアドインを使用できるようにすることをお勧めします。 _すべて_の Office ホストおよびプラットフォームでアドインを使用できるようにするには、 `Requirements`要素の代わりにランタイムチェックを使用します。

次のものをサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。

-  `TableBindings`要件セット。最小バージョンは "1.1" です。

-  `OOXML`要件セット。最小バージョンは "1.1" です。

-  `Document.getSelectedDataAsync`手段.

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

- 要素`Requirements`には`Sets`および`Methods`子要素が含まれています。

- 要素`Sets`には、1つ以上`Set`の要素を含めることができます。 `DefaultMinVersion`すべての子`MinVersion` `Set`要素の既定値を指定します。

- 要素`Set`は、Office ホストがアドインを実行するためにサポートする必要がある要件セットを指定します。 属性`Name`は、要件セットの名前を指定します。 は`MinVersion` 、要件セットの最小バージョンを指定します。 `MinVersion`API メンバーが属する`DefaultMinVersion`要件セットと要件セットのバージョンの詳細については、「 [Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。

- 要素`Methods`には、1つ以上`Method`の要素を含めることができます。Outlook アドインで`Methods`要素を使用することはできません。

- 要素`Method`は、アドインを実行する Office ホストでサポートする必要がある個別のメソッドを指定します。この`Name`属性は必須で、その親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="use-runtime-checks-in-your-javascript-code"></a>JavaScript コードでランタイム チェックを使用する

特定の要件セットが Office ホストでサポートされる場合、追加の機能を提供すると効果的な場合があります。 たとえば、アドインで Word 2016 を実行する場合、既存のアドインで Word JavaScript API を使用することがあります。 その場合、要件セットの名前を指定し、[isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドを使用します。 `isSetSupported`実行時に、アドインを実行している Office ホストが要件セットをサポートするかどうかを指定します。 要件セットがサポートされて`isSetSupported`いる場合は、 **true**を返し、その要件セットの API メンバーを使用する追加のコードを実行します。 Office ホストが要件セットをサポートしてい`isSetSupported`ない場合は、 **false**を返し、追加のコードは実行されません。 次のコードは、で`isSetSupported`使用する構文を示しています。

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- _RequirementSetName_ (必須) は、要件セットの名前を表す文字列です (例: "**ExcelApi**"、"**Mailbox**" など)。 利用できる要件セットの詳細については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」を参照してください。
- _MinimumVersion_ (省略可能) では、`if` ステートメントの範囲内でコードを実行するために、ホストがサポートする必要がある最小要件セットのバージョンを指定します (例: "**1.9**")。

> [!WARNING]
> `isSetSupported`メソッドを呼び出す場合、 `MinimumVersion`パラメーター (指定されている場合) の値は文字列である必要があります。 これは、JavaScript パーサーでは、1.1 や 1.10 のような数値の間の差異を区別できないが、"1.1" や "1.10" などの文字列値ではできるからです。
> `number` のオーバーロードは非推奨になります。

Office `isSetSupported`ホストと`RequirementSetName`関連付けられているを次のように使用します。

|Office ホスト|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|Word|WordApi|

これら`isSetSupported`のホストのメソッドと要件セットは、CDN の最新の Office .js ファイルで使用できます。 CDN から Office .js を使用しない場合は、例外が発生する可能性があるため`isSetSupported` 、アドインが例外を生成することがあります。 詳細については、「[最新の Office JAVASCRIPT API ライブラリを指定する](#specify-the-latest-office-javascript-api-library)」を参照してください。

次のコードの例は、さまざまな要件セットや API メンバーをサポートするさまざまな Office ホストにおいて、アドインで各種の機能を提供する方法を示しています。

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
    // Run additional code when the Office host is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>要件セットにないメソッドを使用したランタイム チェック

API の一部のメンバーは、要件のセットに属していません。 これは、 [Office JavaScript api](../reference/javascript-api-for-office.md)名前空間 ( `Office.` [Outlook メールボックス api](/javascript/api/outlook)以外のすべて) に属する api メンバーではなく、 [Word javascript api](../reference/overview/word-add-ins-reference-overview.md) (すべて`Word.`のもの)、 [Excel javascript api](../reference/overview/excel-add-ins-reference-overview.md) `Excel.`(すべての場合)、または[OneNote javascript api](../reference/overview/onenote-add-ins-javascript-reference.md) ( `OneNote.`あらゆる場合) の名前空間に含まれる api メンバーにのみ適用されます。 要件セットに属さないメソッドにアドインが依存するとき、ランタイム チェックを利用し、メソッドが Office ホストでサポートされているかどうかを判断できます。たとえば、次のコード例のようになります。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例では、ホストが`document.setSelectedDataAsync`サポートしているかどうかを確認します。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](add-in-manifests.md)
- [Office アドインの要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
