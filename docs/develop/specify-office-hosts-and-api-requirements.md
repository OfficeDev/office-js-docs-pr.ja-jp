---
title: Office のホストと API の要件を指定する
description: アドインが期待通Office動作するアプリケーションと API 要件を指定する方法について説明します。
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7e43aa05d543eb55f10c6e700b5011733792a401
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496804"
---
# <a name="specify-office-applications-and-api-requirements"></a>Office アプリケーションと API 要件を指定する

アドインOfficeは、特定の Office アプリケーション (Office ホストとも呼ばれる) または Office JavaScript API (office.js) の特定のメンバーによって異なる場合があります。 たとえば、次のようなアドインがあります。

- 1 つの Office アプリケーション (例: Word または Excel)、またはいくつかのアプリケーションで実行します。
- 一部のバージョンOfficeでのみ使用できる JavaScript API を使用Office。 たとえば、1 回の購入バージョンExcel 2016 JavaScript ライブラリ内Excel関連する API Officeサポートされていません。

このような場合は、アドインが実行できない Office アプリケーションまたは Office バージョンにインストールされていないことを確認する必要があります。

また、ユーザーのアプリケーションとバージョンに基づいて、ユーザーに表示されるアドインの機能を制御OfficeシナリオOfficeがあります。 次の 2 つの例を示します。

- アドインには、テキスト操作など、Word と PowerPoint の両方で役立つ機能がありますが、スライド管理機能など、PowerPoint でしか意味をなさない機能がいくつか追加されています。 アドインが Word で実行されているPowerPointの機能を非表示にする必要があります。
- アドインには、サブスクリプション Excel など、Office アプリケーションの一部のバージョンでサポートされている Office JavaScript API メソッドが必要ですが、一時購入 Excel 2016 など、他のバージョンではサポートされていない機能があります。 ただし、アドインには、その他の機能が含Office JavaScript API メソッドのみを必要とExcel 2016。 このシナリオでは、アドインを Excel 2016 にインストールできる必要がありますが、サポートされていない方法を必要とする機能は、Excel 2016 のユーザーから非表示にしてください。

この記事は、アドインが期待どおりに機能し、できるだけ多くのユーザーが利用できるようにするために選択する必要のあるオプションについて理解するのに役立ちます。

> [!NOTE]
> Office アドインが現在サポートされている場所の詳細なビューについては、「[Office](/javascript/api/requirement-sets) クライアント アプリケーションと Office アドインのプラットフォームの可用性」ページを参照してください。

> [!TIP]
> この記事で説明するタスクの多くは、Office アドインの [Yeoman](yeoman-generator-overview.md) ジェネレーターや Visual Studio の Office アドイン テンプレートの 1 つなど、ツールを使用してアドイン プロジェクトを作成するときに、全体または一部で実行されます。 このような場合は、タスクが実行されたことを確認する必要があるという意味として、タスクを解釈してください。

## <a name="use-the-latest-office-javascript-api-library"></a>JavaScript API ライブラリOfficeを使用する

アドインは、コンテンツ配信ネットワーク (Office) から JavaScript API ライブラリの最新バージョンを読み込CDN。 これを行うには、アドインが開く `script` 最初の HTML ファイルに次のタグが含まれる必要があります。 CDN URL で `/1/` を使用することで、Office.js の最新版が参照されます。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>アドインをOfficeできるアプリケーションを指定する

既定では、指定したアドインの種類 (メール、作業ウィンドウ、またはコンテンツ) でサポートされているすべての Office アプリケーションにアドインをインストールできます。 たとえば、作業ウィンドウ アドインは Access、Excel、OneNote、PowerPoint、Project、および Word で既定でインストールできます。 

アドインがアプリケーションのサブセットにインストール可能Office、マニフェストの [Hosts](/javascript/api/manifest/hosts) 要素と [Host](/javascript/api/manifest/host) 要素を使用します。

たとえば、次の **Hosts** および **Host** 宣言は、アドインが Excel の任意のリリース (Excel on the web、Windows、および iPad を含む)にインストールできますが、他の Office アプリケーションにはインストールできないことを指定します。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

**Hosts** 要素には 1 つ以上の **Host** 要素を含めることができます。 アドインを **インストール可能に** Officeホスト アプリケーションごとに個別の Host 要素が必要です。 属性 `Name` は必須であり、次のいずれかの値に設定できます。

| 名前          | Office クライアント アプリケーション                     | 使用可能なアドインの種類 |
|:--------------|:-----------------------------------------------|:-----------------------|
| データベース      | Access Web アプリ                                | 作業ウィンドウ              |
| Document      | Word on the web、Windows、Mac、iPad            | 作業ウィンドウ              |
| Mailbox       | Outlook on the web、Windows、Mac、Android、iOS | メール                   |
| Notebook      | OneNote on the web                             | 作業ウィンドウ、コンテンツ     |
| Presentation  | PowerPoint on the web、Windows、Mac、iPad      | 作業ウィンドウ、コンテンツ     |
| Project       | Windows での Project                             | 作業ウィンドウ              |
| Workbook      | Excel on the web、Windows、Mac、iPad           | 作業ウィンドウ、コンテンツ     |

> [!NOTE]
> Officeアプリケーションはさまざまなプラットフォームでサポートされ、デスクトップ、Web ブラウザー、タブレット、モバイル デバイスで実行されます。 通常、アドインの実行に使用できるプラットフォームを指定できません。 たとえば、指定した場合`Workbook`は、Excel on the webとWindowsアドインの実行に使用できます。 ただし、指定した場合`Mailbox`、モバイル拡張ポイントを定義しない限り、Outlookクライアントでアドイン[が実行されません](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface)。

> [!NOTE]
> アドイン マニフェストを複数の種類 (メール、作業ウィンドウ、コンテンツ) に適用できない。 つまり、Outlook と他の Office アプリケーションの 1 つでアドインをインストールする場合は、メールの種類マニフェストを持つ 2 つのアドインと、作業ウィンドウマニフェストまたはコンテンツ タイプ マニフェストを持つアドインを 2 つ作成する必要があります。

> [!IMPORTANT]
> SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。 代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>アドインをOfficeできるバージョンとプラットフォームを指定する

Office のバージョンとビルド、またはアドインをインストール可能にするプラットフォームを明示的に指定できません。また、アドインで使用するアドイン機能のサポートが新しいバージョンまたはプラットフォームに拡張されるたびにマニフェストを変更する必要が生じないので、不要です。 代わりに、アドインで必要な API をマニフェストで指定します。 Office API をサポートしない Office バージョンとプラットフォームの組み合わせにアドインがインストールされるのを防ぎ、アドインが [マイ アドイン] に表示されないのを確認します。

> [!IMPORTANT]
> 基本マニフェストのみを使用して、アドインが重要な値である必要がある API メンバーを指定します。 アドインが一部の機能に API を使用しているが、API を必要としないその他の便利な機能がある場合は、API をサポートしないプラットフォームと Office バージョンの組み合わせでインストール可能で、それらの組み合わせに対するエクスペリエンスが低下するアドインを設計する必要があります。 詳細については、「代替エクスペリエンスを [設計する」を参照してください](#design-for-alternate-experiences)。

### <a name="requirement-sets"></a>要件セット

アドインに必要な API を指定するプロセスを簡略化するために、Office API を要件セットで *グループ化します*。 共通 API オブジェクト モデル [の API](understanding-the-javascript-api-for-office.md#api-models) は、サポートする開発機能によってグループ化されます。 たとえば、テーブル バインドに接続されている API はすべて、"TableBindings 1.1" と呼ばれる要件セットに含されます。 Application 固有のオブジェクト モデル [の API](understanding-the-javascript-api-for-office.md#api-models) は、実稼働アドインで使用するためにリリースされた場合にグループ化されます。

要件セットはバージョン管理されています。 たとえば、ダイアログ ボックスをサポートする API [は](../design/dialog-boxes.md) 、要件セット DialogApi 1.1 にあるとします。 作業ウィンドウからダイアログへのメッセージングを有効にする追加の API がリリースされると、DialogApi 1.1 のすべての API と共に DialogApi 1.2 にグループ化されました。 *要件セットの各バージョンは、以前のすべてのバージョンのスーパーセットです。*

要件セットのサポートは、Officeアプリケーション、Officeアプリケーションのバージョン、およびアプリケーションが実行されているプラットフォームによって異なります。 たとえば、DialogApi 1.2 は Office 2021 より前の 1 回購入バージョンの Office ではサポートされていませんが、DialogApi 1.1 は Office 2013 に戻るすべてのワンタイム購入バージョンでサポートされています。 使用する API をサポートするプラットフォームと Office バージョンのすべての組み合わせでアドインをインストール可能にし、アドインに必要な各要件セットの最小バージョンをマニフェストで常に指定する必要があります。 これを行う方法の詳細については、この記事の後半で説明します。

> [!TIP]
> 要件セットのバージョン管理の詳細については、「[Office](office-versions-and-requirement-sets.md#office-requirement-sets-availability) 要件セットの可用性」、および各 API に関する要件セットと情報の完全な一覧については、「[Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets) アドイン要件セット」を参照してください。 ほとんどの API の参照トピックOffice.js、その API に属する要件セット (必要な場合) も指定します。

> [!NOTE]
> 一部の要件セットには、マニフェスト要素も関連付けられているものがあります。 この [事実がアドインデザインに関連する](#specify-requirements-in-a-versionoverrides-element) 場合の詳細については、「VersionOverrides 要素での要件の指定」を参照してください。

#### <a name="apis-not-in-a-requirement-set"></a>API が要件セットに含めされていない

アプリケーション固有のモデル内のすべての API は要件セット内ですが、共通 API モデル内の API の一部は要件セットではありません。 また、アドインで必要な場合に、マニフェストでこれらの設定されていない API のいずれかを指定する方法も用意されています。 詳細については、この記事で後述します。

### <a name="requirements-element"></a>Requirements 要素

Requirements 要素とその子要素 [Sets](/javascript/api/manifest/sets) と [Methods](/javascript/api/manifest/requirements) を[](/javascript/api/manifest/methods)使用して、アドインをインストールするために Office アプリケーションでサポートする必要がある最小要件セットまたは API メンバーを指定します。 

Office アプリケーションまたはプラットフォームが **Requirements** 要素で指定された要件セットまたは API メンバーをサポートしない場合、アドインは、そのアプリケーションまたはプラットフォームでは実行されません。また、My アドインには表示されません。

> [!NOTE]
> **Requirements 要素** は、すべてのアドインに対して省略可能です(ただし、Outlookを除く)。ルート要素`xsi:type`の属性が存在`OfficeApp``MailBox`する場合は、アドインで必要な MailBox 要件セットの最小バージョンを指定する **Requirements** 要素が必要です。 詳細については、「[JavaScript API 要件セットOutlook参照してください](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)。

次のコード例は、以下をサポートするすべてのアプリケーションでインストール可能なアドインOffice示しています。

-  `TableBindings` 要件セット 。最小バージョンは "1.1" です。
-  `OOXML` 要件セット 。最小バージョンは "1.1" です。
-  `Document.getSelectedDataAsync` メソッド。

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

- **Requirements 要素には**、Sets と **Methods** **の子要素が** 含まれます。
- **Sets 要素には**、1 つ以上の **Set 要素を含** めできます。 `DefaultMinVersion` すべての子 Set 要素の `MinVersion` 既定値を **指定** します。
- [Set 要素](/javascript/api/manifest/set)は、アドインをインストール可能Officeアプリケーションがサポートする必要がある要件セットを指定します。 属性 `Name` は、要件セットの名前を指定します。 要件 `MinVersion` セットの最小バージョンを指定します。 `MinVersion` 親セットの属性の値 `DefaultMinVersion` を上書き **します**。
- **Methods 要素には**、1 つ以上の [Method 要素を含](/javascript/api/manifest/method)めできます。 Outlook アドインで **Methods** 要素を使用することはできません。
- **Method 要素** は、アドインをインストール可能Officeアプリケーションがサポートする必要がある個別のメソッドを指定します。 属性 `Name` は必須であり、親オブジェクトで修飾されたメソッドの名前を指定します。

## <a name="design-for-alternate-experiences"></a>代替エクスペリエンスの設計

アドイン プラットフォームが提供する機能Office機能は、次の 3 種類に分けて役立ちます。

- アドインがインストールされた直後に使用できる機能拡張。 この種の機能を利用するには、マニフェストで [VersionOverrides](/javascript/api/manifest/versionoverrides) 要素を構成します。 この種の機能の例は、 [カスタム](../design/add-in-commands.md) リボン ボタンとメニューであるアドイン コマンドです。
- アドインが実行され、JavaScript API で実装されている場合にのみ使用可能な機能Office.js機能。たとえば、[ [ダイアログ ボックス] などです](../design/dialog-boxes.md)。
- 実行時にのみ使用できますが、 **VersionOverrides** 要素の JavaScript と構成Office.js組み合わせて実装されている機能。 これらの例には、Excel[、](../excel/custom-functions-overview.md)シングル サインオン、カスタム コンテキスト [タブがあります](../design/contextual-tabs.md)。 [](sso-in-office-add-ins.md)

アドインが一部の機能に特定の機能拡張機能を使用しているが、拡張機能を必要としないその他の便利な機能がある場合は、拡張機能をサポートしないプラットフォームと Office バージョンの組み合わせでアドインをインストール可能に設計する必要があります。 これらの組み合わせに関する貴重なエクスペリエンスを提供できます。 

この設計は、機能拡張機能の実装方法に応じて異なる方法で実装します。 

- JavaScript で完全に実装された機能については、「ランタイム がメソッドと要件セットの [サポートをチェックする」を参照してください](#runtime-checks-for-method-and-requirement-set-support)。
- VersionOverrides 要素を構成する必要がある機能については、「 **VersionOverrides** 要素での要件の指定 [」を参照してください](#specify-requirements-in-a-versionoverrides-element)。

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>ランタイムがメソッドと要件セットのサポートをチェックする 

実行時にテストを行い、[isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) メソッドOfficeをサポートするかどうかを確認します。 要件セットの名前と最小バージョンをパラメーターとして渡します。 要件セットがサポートされている場合は、 `isSetSupported` true を返 **します**。 次のコードは一例です。

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is one-time purchase Word 2013 (which does not support WordApi 1.1).
}
```
このコードについては、以下の点に注意してください。

- 最初のパラメーターは必須です。 これは、要件セットの名前を表す文字列です。 利用できる要件セットの詳細については、「[Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)」を参照してください。
- 2 番目のパラメーターは省略可能です。 これは、ステートメント内のコードを実行するために Office `if` アプリケーションがサポートする必要がある最小要件セットのバージョン ("**1.9**" など) を指定する文字列です。 使用しない場合は、バージョン "1.1" が想定されます。

> [!WARNING]
> メソッドを呼び出 `isSetSupported` す場合、2 番目のパラメーターの値 (指定されている場合) は、数値ではなく文字列である必要があります。 JavaScript パーサーは、1.1 や 1.10 などの数値を区別することはできませんが、"1.1" や "1.10" などの文字列値を使用できます。

次の表に、アプリケーション固有の API モデルの要件セット名を示します。

|Office アプリケーション|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|PowerPoint|PowerPointApi|
|Word|WordApi|

次に、共通 API モデル要件セットの 1 つでメソッドを使用する例を示します。

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Word doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE] 
> これらの`isSetSupported`アプリケーションのメソッドと要件セットは、アプリケーションのOffice.jsファイルで使用CDN。 `isSetSupported` CDN から Office.js を使用しない場合、未定義の古いバージョンのライブラリを使用している場合、アドインによって例外が生成される場合があります。 詳細については、「[Use the latest Office JavaScript API ライブラリ」を参照してください](#use-the-latest-office-javascript-api-library)。

アドインが要件セットの一部ではないメソッドに依存している場合は、ランタイム チェックを使用して、次のコード例に示すように、メソッドが Office アプリケーションでサポートされているかどうかを判断します。 要件セットに属さないメソッドの詳細な一覧については、「[Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set)」を参照してください。

> [!NOTE]
> アドインのコードでのこの種のランタイム チェックは、限定的に使用することをお勧めします。

次のコード例は、アプリケーションがサポートOfficeチェックします`document.setSelectedDataAsync`。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>VersionOverrides 要素で要件を指定する

[VersionOverrides](/javascript/api/manifest/versionoverrides) 要素は、アドイン コマンド (カスタム リボン ボタンやメニュー) など、アドインのインストール直後に使用できる必要がある機能をサポートするために、主にマニフェスト スキーマに追加されましたが、排他的ではありません。 Officeマニフェストを解析するときに、これらの機能について知っている必要があります。 

アドインでこれらの機能の 1 つを使用しているが、アドインは重要であり、この機能をサポートしない Office バージョンでもインストール可能である必要があるとします。 このシナリオでは、基本要素の子としてではなく **、VersionOverrides** 要素 [](/javascript/api/manifest/sets)自体の子として含める [Requirements](/javascript/api/manifest/requirements) 要素 (および子の Sets と [Methods](/javascript/api/manifest/methods) 要素) を使用して機能を識別`OfficeApp`します。 この場合の効果は、Office はアドインのインストールを許可しますが、Office は機能がサポートされていない Office バージョンの **VersionOverrides** 要素の子要素の一部を無視します。

具体的には、Hosts 要素など、基本マニフェストの要素をオーバーライドする **VersionOverrides** の子要素は無視され、代わりに基本マニフェストの対応する要素が使用されます。 ただし、 **VersionOverrides** には、基本マニフェストの設定を上書きするのではなく、追加の機能を実際に実装する子要素が含まれます。 2 つの例は、 と `WebApplicationInfo` です `EquivalentAddins`。 **VersionOverrides のこれらの部分** は、対応する機能をサポートしているプラットフォームとバージョンOffice無視されません。  

**Requirements** 要素の子孫要素の詳細については、この記事の「[Requirements 要素](#requirements-element)」を参照してください。

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
>  **VersionOverrides** で **Requirements** 要素を使用する前に、要件をサポートしないプラットフォームとバージョンの組み合わせでは、要件を必要としない機能を呼び出すアドイン コマンドもインストールされません。 たとえば、2 つのカスタム リボン ボタンを持つアドインを検討します。 そのうちの 1 つはOffice **ExcelApi 1.4** (以降) で使用できる JavaScript API を呼び出します。 その他の呼び出し API は **、ExcelApi 1.9** (以降) でのみ使用できます。 **VersionOverrides** に **ExcelApi 1.9** の要件を設定すると、1.9 がサポートされていない場合、どちらのボタンもリボンに表示されません。 このシナリオのより良い戦略は、「メソッドと要件セットのサポートをランタイム チェックする」で説明されている [手法を使用する方法です](#runtime-checks-for-method-and-requirement-set-support)。 2 番目のボタンによって呼び出されるコードは、`isSetSupported`**まず ExcelApi 1.9 のサポートを確認するために使用します**。 サポートされていない場合、このコードは、アドインのこの機能がバージョンのアドインで使用できないというメッセージをユーザーにOffice。 

> [!TIP]
> 基本マニフェストに既に表示されている **VersionOverrides で Requirement** 要素を繰り返す意味はありません。 基本マニフェストで要件が指定されている場合、アドインは要件がサポートされていない場所にインストールできないので、Office は **VersionOverrides** 要素を解析する必要さえも持てない。 

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](add-in-manifests.md)
- [Office アドインの要件セット](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
