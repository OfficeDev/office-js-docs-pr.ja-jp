---
title: Office のホストと API の要件を指定する
description: アドインが想定どおりに動作するように Office アプリケーションと API 要件を指定する方法について説明します。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60d69c9fae136e73bf9920c7dc96f13d832331fd
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810296"
---
# <a name="specify-office-applications-and-api-requirements"></a>Office アプリケーションと API 要件を指定する

Office アドインは、特定の Office アプリケーション (Office ホストとも呼ばれます) または Office JavaScript API (office.js) の特定のメンバーに依存する場合があります。 たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。
- 一部のバージョンの Office でのみ使用できる Office JavaScript API を使用します。 たとえば、ボリューム ライセンスの永続的なバージョンのExcel 2016では、Office JavaScript ライブラリ内のすべての Excel 関連 API がサポートされているわけではありません。

このような状況では、アドインが実行できない Office アプリケーションまたは Office バージョンにインストールされないようにする必要があります。

また、Office アプリケーションと Office のバージョンに基づいて、アドインのどの機能をユーザーに表示するかを制御するシナリオもあります。 次の 2 つの例を示します。

- アドインには、テキスト操作など、Word と PowerPoint の両方で役立つ機能がありますが、スライド管理機能など、PowerPoint でのみ意味を持ついくつかの追加機能があります。 アドインが Word で実行されている場合は、PowerPoint のみの機能を非表示にする必要があります。
- アドインには、Office アプリケーションの一部のバージョン (Microsoft 365 サブスクリプション Excel など) でサポートされているが、ボリューム ライセンスの永続的なExcel 2016など、他のバージョンではサポートされていない Office JavaScript API メソッドが必要な機能があります。 ただし、アドインには、ボリューム ライセンスの永続的なExcel 2016でサポート *されている* Office JavaScript API メソッドのみを必要とする他の機能があります。 このシナリオでは、そのバージョンのExcel 2016にアドインをインストールできる必要がありますが、サポートされていない方法を必要とする機能は、それらのユーザーに表示しないようにする必要があります。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> Office アドインが現在サポートされている場所の概要については、「Office アドインの [Office クライアント アプリケーションとプラットフォームの可用性](/javascript/api/requirement-sets) 」ページを参照してください。

> [!TIP]
> この記事で説明するタスクの多くは、Office アドイン [の Yeoman ジェネレーター](yeoman-generator-overview.md) や Visual Studio の Office アドイン テンプレートの 1 つなど、ツールを使用してアドイン プロジェクトを作成するときに、全体または一部で実行されます。 このような場合は、タスクが完了したことを確認する必要があることを意味として解釈してください。

## <a name="use-the-latest-office-javascript-api-library"></a>最新の Office JavaScript API ライブラリを使用する

アドインは、コンテンツ配信ネットワーク (CDN) から最新バージョンの Office JavaScript API ライブラリを読み込む必要があります。 これを行うには、アドインが開く最初の HTML ファイルに次 `script` のタグがあることを確認します。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>アドインをホストできる Office アプリケーションを指定する

既定では、アドインは、指定したアドインの種類 (つまり、メール、作業ウィンドウ、またはコンテンツ) でサポートされているすべての Office アプリケーションにインストールできます。 たとえば、作業ウィンドウ アドインは、Access、Excel、OneNote、PowerPoint、Project、Word に既定でインストールできます。

アドインを Office アプリケーションのサブセットに確実にインストールできるようにするには、マニフェストで [Hosts](/javascript/api/manifest/hosts) 要素と [Host](/javascript/api/manifest/host) 要素を使用します。

たとえば、次 **\<Hosts\>** の宣言と **\<Host\>** 宣言では、アドインが Excel の任意のリリースにインストールできることを指定します。これには、Excel on the web、Windows、iPad が含まれますが、他の Office アプリケーションにはインストールできません。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

要素には **\<Hosts\>** 、1 つ以上の **\<Host\>** 要素を含めることができます。 アドインをインストールできる Office アプリケーションごとに個別 **\<Host\>** の要素が必要です。 `Name`属性は必須であり、次のいずれかの値に設定できます。

| 名前          | Office クライアント アプリケーション                     | 使用可能なアドインの種類 |
|:--------------|:-----------------------------------------------|:-----------------------|
| データベース      | Access Web アプリ                                | 作業ウィンドウ              |
| ドキュメント      | Word on the web, Windows, Mac, iPad            | 作業ウィンドウ              |
| Mailbox       | Outlook on the web、Windows、Mac、Android、iOS | メール                   |
| Notebook      | OneNote on the web                             | 作業ウィンドウ、コンテンツ     |
| Presentation  | PowerPoint on the web、Windows、Mac、iPad      | 作業ウィンドウ、コンテンツ     |
| Project       | Windows での Project                             | 作業ウィンドウ              |
| ブック      | Excel on the web、Windows、Mac、iPad           | 作業ウィンドウ、コンテンツ     |

> [!NOTE]
> Office アプリケーションは、さまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行されます。 通常、アドインの実行に使用できるプラットフォームを指定することはできません。 たとえば、 を指定`Workbook`した場合、Excel on the webと Windows の両方を使用してアドインを実行できます。 ただし、 を指定 `Mailbox`した場合、 [モバイル拡張機能ポイント](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface)を定義しない限り、アドインは Outlook モバイル クライアントでは実行されません。

> [!NOTE]
> アドイン マニフェストを複数の種類 (メール、作業ウィンドウ、またはコンテンツ) に適用することはできません。 つまり、アドインを Outlook と他の Office アプリケーションのいずれかにインストールできるようにする場合は、 *2 つの* アドイン (1 つはメール型マニフェスト、もう 1 つは作業ウィンドウまたはコンテンツ タイプ マニフェスト) を作成する必要があります。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>アドインをホストできる Office のバージョンとプラットフォームを指定する

Office のバージョンとビルド、またはアドインをインストール可能にするプラットフォームを明示的に指定することはできません。また、アドインで使用するアドイン機能のサポートが新しいバージョンまたはプラットフォームに拡張されるたびにマニフェストを変更する必要があるため、必要ありません。 代わりに、アドインに必要な API をマニフェストで指定します。 Office は、API をサポートしていない Office バージョンとプラットフォームの組み合わせにアドインがインストールされないようにし、アドインが **マイ アドイン** に表示されないようにします。

> [!IMPORTANT]
> ベース マニフェストを使用して、アドインが重要な値である必要がある API メンバーを指定するだけです。 アドインで一部の機能に API を使用しているが、API を必要としないその他の便利な機能がある場合は、API をサポートしていないが、それらの組み合わせでエクスペリエンスが低下するプラットフォームと Office のバージョンの組み合わせにインストールできるようにアドインを設計する必要があります。 詳細については、「 [代替エクスペリエンスの設計](#design-for-alternate-experiences)」を参照してください。

### <a name="requirement-sets"></a>要件セット

アドインで必要な API を指定するプロセスを簡略化するために、Office はほとんどの API を *要件セットにまとめます*。 [共通 API オブジェクト モデルの API](understanding-the-javascript-api-for-office.md#api-models) は、サポートされている開発機能によってグループ化されます。 たとえば、テーブル バインドに接続されているすべての API は、"TableBindings 1.1" という要件セットにあります。 [アプリケーション固有のオブジェクト モデル](understanding-the-javascript-api-for-office.md#api-models)の API は、運用環境のアドインで使用するためにリリースされたときによってグループ化されます。

要件セットはバージョン管理されます。 たとえば、 [ダイアログ ボックス](../develop/dialog-api-in-office-add-ins.md) をサポートする API は、要件セット DialogApi 1.1 にあります。 作業ウィンドウからダイアログへのメッセージングを有効にする追加の API がリリースされると、DialogApi 1.1 のすべての API と共に DialogApi 1.2 にグループ化されました。 *要件セットの各バージョンは、以前のすべてのバージョンのスーパーセットです。*

要件セットのサポートは、Office アプリケーション、Office アプリケーションのバージョン、および実行されているプラットフォームによって異なります。 たとえば、DialogApi 1.2 は、Office 2021前のボリューム ライセンスの永続的なバージョンの Office ではサポートされていませんが、DialogApi 1.1 は Office 2013 に戻るすべての永続的なバージョンでサポートされています。 使用する API をサポートするプラットフォームと Office バージョンのすべての組み合わせにアドインをインストールできるようにする必要があるため、必ずマニフェストで、アドインに必要な各要件セットの *最小* バージョンを指定する必要があります。 これを行う方法の詳細については、この記事の後半で説明します。

> [!TIP]
> 要件セットのバージョン管理の詳細については、「 [Office 要件セットの可用性](office-versions-and-requirement-sets.md#office-requirement-sets-availability)」を参照し、要件セットの完全な一覧と各 API に関する情報については、 [Office アドイン要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)から始めます。 ほとんどのOffice.js API のリファレンス トピックでは、所属する要件セット (存在する場合) も指定します。

> [!NOTE]
> 一部の要件セットには、マニフェスト要素も関連付けられています。 この事実がアドインの設計に関連する場合の詳細については、「 [VersionOverrides 要素での要件の指定](#specify-requirements-in-a-versionoverrides-element) 」を参照してください。

#### <a name="apis-not-in-a-requirement-set"></a>要件セットに含まれていない API

アプリケーション固有のモデル内のすべての API は要件セットに含まれていますが、Common API モデルの API の一部は含まれていません。 また、アドインに必要な場合に、マニフェストでこれらのセットレス API のいずれかを指定する方法もあります。 詳細については、この記事で後述します。

### <a name="requirements-element"></a>Requirements 要素

[Requirements](/javascript/api/manifest/requirements) 要素とその子要素 [Sets](/javascript/api/manifest/sets) と [Methods](/javascript/api/manifest/methods) を使用して、アドインをインストールするために Office アプリケーションでサポートする必要がある最小要件セットまたは API メンバーを指定します。

Office アプリケーションまたはプラットフォームが 要素で指定された要件セットまたは API メンバーを **\<Requirements\>** サポートしていない場合、アドインはそのアプリケーションまたはプラットフォームで実行されないため、 **マイ アドイン** には表示されません。

> [!NOTE]
> 要素は **\<Requirements\>**、Outlook アドインを除くすべてのアドインでは省略可能です。ルート要素の`xsi:type`属性が である場合は、`MailApp`アドインに必要な **\<Requirements\>** メールボックス要件セットの最小バージョンを指定する要素が必要`OfficeApp`です。 詳細については、「 [Outlook JavaScript API 要件セット](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)」を参照してください。

次のコード例は、次をサポートするすべての Office アプリケーションにインストール可能なアドインを構成する方法を示しています。

- `TableBindings` 要件セット。最小バージョンは "1.1" です。
- `OOXML` 要件セット。最小バージョンは "1.1" です。
- `Document.getSelectedDataAsync` メソッド。

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```

この例については、次の点に注意してください。

- 要素には **\<Requirements\>**、 要素と **\<Methods\>** 子要素が **\<Sets\>** 含まれています。
- 要素には **\<Sets\>** 、1 つ以上の **\<Set\>** 要素を含めることができます。 `DefaultMinVersion`は、すべての子 **\<Set\>** 要素の既定値を指定します`MinVersion`。
- [Set](/javascript/api/manifest/set) 要素は、アドインをインストールできるようにするために Office アプリケーションがサポートする必要がある要件セットを指定します。 属性は `Name` 、要件セットの名前を指定します。 は `MinVersion` 、要件セットの最小バージョンを指定します。 `MinVersion`は、親 **\<Sets\>** の 属性の`DefaultMinVersion`値をオーバーライドします。
- 要素には **\<Methods\>** 、1 つ以上の [Method 要素を](/javascript/api/manifest/method) 含めることができます。 Outlook アドインで 要素を **\<Methods\>** 使用することはできません。
- 要素は **\<Method\>** 、アドインをインストールできるようにするために Office アプリケーションがサポートする必要がある個々のメソッドを指定します。 `Name`属性は必須であり、その親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="design-for-alternate-experiences"></a>代替エクスペリエンスの設計

Office アドイン プラットフォームが提供する機能拡張機能は、次の 3 種類に分けることができます。

- アドインのインストール直後に使用できる機能拡張機能。 この種の機能を利用するには、マニフェストで [VersionOverrides](/javascript/api/manifest/versionoverrides) 要素を構成します。 この種の機能の例としては、カスタム リボン ボタンとメニューである [アドイン コマンド](../design/add-in-commands.md)があります。
- アドインが実行されていて、Office.js JavaScript API で実装されている場合にのみ使用できる機能拡張機能。たとえば、 [ダイアログ ボックス](../develop/dialog-api-in-office-add-ins.md)です。
- 実行時にのみ使用できますが、Office.js JavaScript と要素内の構成の組み合わせで **\<VersionOverrides\>** 実装される機能拡張機能。 これらの例としては、 [Excel カスタム関数](../excel/custom-functions-overview.md)、 [シングル サインオン](sso-in-office-add-ins.md)、 [カスタム コンテキスト タブなどがあります](../design/contextual-tabs.md)。

アドインで機能の一部に特定の機能拡張機能を使用しているが、機能拡張機能を必要としないその他の便利な機能がある場合は、拡張機能機能をサポートしていないプラットフォームと Office のバージョンの組み合わせにインストールできるようにアドインを設計する必要があります。 これらの組み合わせに関する貴重な、しかし減少した経験を提供することができます。

この設計は、機能拡張機能の実装方法に応じて異なる方法で実装します。

- JavaScript で完全に実装される機能については、 [メソッドと要件セットのサポートに関するランタイム チェックに関する](#runtime-checks-for-method-and-requirement-set-support)ページを参照してください。
- 要素を構成 **\<VersionOverrides\>** する必要がある機能については、「 [VersionOverrides 要素での要件の指定](#specify-requirements-in-a-versionoverrides-element)」を参照してください。

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>メソッドと要件セットのサポートのランタイム チェック

実行時にテストして、ユーザーの Office が [isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) メソッドで要件セットをサポートしているかどうかを検出します。 要件セットの名前と最小バージョンをパラメーターとして渡します。 要件セットがサポートされている場合は、 `isSetSupported` を返します `true`。 次のコードは一例です。

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is perpetual Word 2013 (which does not support WordApi 1.1).
}
```

このコードについては、以下の点に注意してください。

- 最初のパラメーターが必要です。 これは、要件セットの名前を表す文字列です。 利用できる要件セットの詳細については、「[Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)」を参照してください。
- 2 番目のパラメーターは省略可能です。 これは、ステートメント内 `if` のコードを実行するために Office アプリケーションがサポートする必要がある最小要件セット のバージョンを指定する文字列です (例: "**1.9**")。 使用しない場合は、バージョン "1.1" が想定されます。

> [!WARNING]
> メソッドを `isSetSupported` 呼び出すとき、2 番目のパラメーターの値 (指定した場合) は数値ではなく文字列にする必要があります。 JavaScript パーサーは、1.1 や 1.10 などの数値を区別できません。一方、"1.1" や "1.10" などの文字列値の場合は区別できません。

次の表は、アプリケーション固有の API モデルの要件セット名を示しています。

|Office アプリケーション|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|PowerPoint|PowerPointApi|
|Word|WordApi|

次に、共通 API モデル要件セットのいずれかで メソッドを使用する例を示します。

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Office application doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE]
> これらのアプリケーションのメソッドと要件セットは `isSetSupported` 、CDN の最新のOffice.js ファイルで使用できます。 CDN からOffice.jsを使用しない場合、未定義のライブラリの古いバージョンを使用している場合、アドインによって `isSetSupported` 例外が生成される可能性があります。 詳細については、「 [最新の Office JavaScript API ライブラリを使用する](#use-the-latest-office-javascript-api-library)」を参照してください。

アドインが要件セットの一部ではないメソッドに依存している場合は、次のコード例に示すように、ランタイム チェックを使用して、メソッドが Office アプリケーションでサポートされているかどうかを判断します。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例では、Office アプリケーションが をサポートしているかどうかを確認します `document.setSelectedDataAsync`。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>VersionOverrides 要素で要件を指定する

[VersionOverrides](/javascript/api/manifest/versionoverrides) 要素は、アドイン コマンド (カスタム リボン ボタンとメニュー) など、アドインのインストール直後に使用できる必要がある機能をサポートするために、主にマニフェスト スキーマに追加されましたが、排他的ではありません。 Office では、アドイン マニフェストを解析するときに、これらの機能について理解しておく必要があります。

アドインでこれらの機能のいずれかを使用しているが、アドインは貴重であり、機能をサポートしていない Office バージョンでもインストール可能であるとします。 このシナリオでは、基本`OfficeApp`要素の子としてではなく、要素自体の **\<VersionOverrides\>** 子として含める [Requirements](/javascript/api/manifest/requirements) 要素 (およびその子 [Sets](/javascript/api/manifest/sets) 要素と [Methods](/javascript/api/manifest/methods) 要素) を使用して特徴を特定します。 これを行うと、Office はアドインのインストールを許可しますが、Office は機能がサポートされていない Office バージョンの要素の **\<VersionOverrides\>** 子要素の特定を無視します。

具体的には、 要素などの **\<Hosts\>** 基本マニフェスト内の **\<VersionOverrides\>** 要素をオーバーライドする の子要素は無視され、代わりにベース マニフェストの対応する要素が使用されます。 ただし、 には、基本マニフェストの **\<VersionOverrides\>** 設定をオーバーライドするのではなく、実際に追加の機能を実装する子要素が存在する可能性があります。 と の 2 つの例を示`WebApplicationInfo``EquivalentAddins`します。 Office の **\<VersionOverrides\>** プラットフォームとバージョンが対応する機能をサポートしていると仮定すると、 のこれらの部分は無視 *されません* 。  

要素の子孫要素 **\<Requirements\>** の詳細については、この記事の「 [Requirements 要素](#requirements-element) 」を参照してください。

次に例を示します。

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> 要件を **\<Requirements\>** サポートしていないプラットフォームとバージョンの組み合わせ *では、要件* を必要としない *機能を呼び出* すアドイン コマンドもインストールされないため、 に要素 **\<VersionOverrides\>** を含める前に細心の注意を払います。 たとえば、2 つのカスタム リボン ボタンを持つアドインを考えてみましょう。 そのうちの 1 つは、要件セット **ExcelApi 1.4** (以降) で使用できる Office JavaScript API を呼び出します。 他の API は、 **ExcelApi 1.9** (以降) でのみ使用できる API を呼び出します。 **に ExcelApi 1.9** の要件を **\<VersionOverrides\>** 設定した場合、1.9 がサポートされていない場合、どちらのボタン *も* リボンに表示されません。 このシナリオでは、「 [メソッドと要件セットのサポートのランタイム チェック](#runtime-checks-for-method-and-requirement-set-support)」で説明されている手法を使用することをお勧めします。 2 番目のボタンによって呼び出されたコードは、最初に **ExcelApi 1.9** のサポートを確認するために使用`isSetSupported`します。 サポートされていない場合は、アドインのこの機能が Office のバージョンでは使用できないというメッセージがユーザーに表示されます。

> [!TIP]
> ベース マニフェストに既に表示されている 内の **\<VersionOverrides\>** **Requirement** 要素を繰り返しても意味がありません。 要件がベース マニフェストで指定されている場合は、要件がサポートされていない場所にアドインをインストールできないため、Office は要素を **\<VersionOverrides\>** 解析することもできません。

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](add-in-manifests.md)
- [Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
