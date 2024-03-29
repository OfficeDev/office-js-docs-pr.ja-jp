---
title: アドイン コマンドを有効または無効にする
description: Office Web アドインのカスタム リボン ボタンとメニュー項目の有効または無効の状態を変更する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 502c9247a6c63775c562dab7479e0ca926f14154
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423049"
---
# <a name="enable-and-disable-add-in-commands"></a>アドイン コマンドを有効または無効にする

アドインの一部の機能を特定のコンテキストでのみ使用可能にする必要がある場合、カスタム アドイン コマンドをプログラムで有効または無効にすることができます。 たとえば、表の見出しを変更する関数は、カーソルが表の中にある場合にのみ有効にする必要があります。

Office クライアント アプリケーションが開いたときに、コマンドを有効にするか無効にするかを指定することもできます。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Office アプリケーションとプラットフォームのサポートのみ

この記事で説明する API は、Excel、PowerPoint、および Word でのみ使用できます。

### <a name="test-for-platform-support-with-requirement-sets"></a>要件セットを使用したプラットフォーム サポートのテスト

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定された要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションとプラットフォームの組み合わせがアドインで必要な API をサポートしているかどうかを判断します。 詳細については、「 [Office のバージョンと要件セット](../develop/office-versions-and-requirement-sets.md)」を参照してください。

有効/無効 API は [、RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) 要件セットに属します。

> [!NOTE]
> **RibbonApi 1.1** 要件セットはマニフェストではまだサポートされていないため、マニフェストの **\<Requirements\>** セクションでは指定できません。 サポートをテストするには、コードが呼び出す `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`必要があります。 その呼び出 *しが* 返された `true`場合にのみ、コードは enable/disable API を呼び出すことができます。 戻`false`り値の`isSetSupported`呼び出しの場合、すべてのカスタム アドイン コマンドが常に有効になります。 **RibbonApi 1.1** 要件セットがサポートされていない場合の動作を考慮するために、運用アドインとアプリ内指示を設計する必要があります。 使用 `isSetSupported`の詳細と例については、「 [Office アプリケーションと API の要件の指定](../develop/specify-office-hosts-and-api-requirements.md)、特に [メソッドと要件セットのサポートに関するランタイム チェック](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)」を参照してください。 (「この記事の [アドインをホストできる Office バージョンとプラットフォームを指定](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) する」セクションは、リボン 1.1 には適用されません)。

## <a name="shared-runtime-required"></a>共有ランタイムが必要

この記事で説明する API とマニフェストマークアップでは、アドインのマニフェストで [共有ランタイム](../testing/runtimes.md#shared-runtime)を使用するように指定する必要があります。 これを行うには、次の手順に従います。

1. マニフェストの [Runtimes](/javascript/api/manifest/runtimes) 要素で、子要素の `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />` を追加します。 (マニフェストに要素がまだ **\<Runtimes\>** ない場合は、セクションの要素の下に最初の **\<Host\>****\<VersionOverrides\>** 子として作成します)。
2. マニフェストの [Resources.Urls](/javascript/api/manifest/resources) セクションで、子要素の `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />` を追加します。ここでは、`{MyDomain}` はアドインのドメインで、`{path-to-start-page}` はアドインの開始ページのパスになります (例: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`)。
3. アドインに作業ウィンドウ、関数ファイル、または Excel カスタム関数が含まれているかどうかに応じて、次の 3 つの手順のうち 1 つ以上を実行する必要があります。

    - アドインに作業ウィンドウが含まれている場合は、アクションの`resid`属性を設定 [します](/javascript/api/manifest/action)。[SourceLocation](/javascript/api/manifest/sourcelocation) 要素は、手順 1 の要素に使用`resid`**\<Runtime\>** したものとまったく同じ文字列です。たとえば、 `Contoso.SharedRuntime.Url`. そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。
    - アドインに Excel カスタム関数が含まれている場合は、Page の属性を`resid`設定 [します](/javascript/api/manifest/page)。[SourceLocation](/javascript/api/manifest/sourcelocation) 要素は、手順 1 の`resid`**\<Runtime\>** 要素に使用したものとまったく同じ文字列です。たとえば、 `Contoso.SharedRuntime.Url`. そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。
    - アドインに関数ファイルが含まれている場合は、[FunctionFile](/javascript/api/manifest/functionfile) 要素の属性を、手順 1 の要素で使用したものと`resid`**\<Runtime\>** まったく同じ文字列に設定`resid`します。たとえば、 `Contoso.SharedRuntime.Url`. そうすると要素は `<FunctionFile resid="Contoso.SharedRuntime.Url"/>` のようになります。

## <a name="set-the-default-state-to-disabled"></a>既定の状態を無効に設定する

既定では、Office アプリケーションの起動時にすべてのアドイン コマンドが有効になります。 Office アプリケーションの起動時にカスタム ボタンまたはメニュー項目を無効にするには、マニフェストで指定します。 コントロールの宣言の [Action](/javascript/api/manifest/action) 要素の *直下* (内部ではない) に、[Enabled](/javascript/api/manifest/enabled) 要素 (値は `false`) を追加するだけで無効にすることができます。 基本的な構造を次に示します。

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>プログラムで状態を変更する

アドイン コマンドの有効な状態を変更するには、以下の手順が重要になります。

1. (1) マニフェストで宣言されている ID でコマンドとその親グループとタブを指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトを作成します。(2) コマンドの有効または無効の状態を指定します。
2. **RibbonUpdaterData** オブジェクトを [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) メソッドに渡します。

次に簡単な例を示します。 "MyButton"、"OfficeAddinTab1"、"CustomGroup111" はマニフェストからコピーされることに注意してください。

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "OfficeAppTab1", 
                groups: [
                    {
                      id: "CustomGroup111",
                      controls: [
                        {
                            id: "MyButton", 
                            enabled: true
                        }
                      ]
                    }
                ]
            }
        ]
    });
}
```

また、**RibbonUpdateData** オブジェクトを簡単に構築できるように、いくつかのインターフェイスも (何種類か) 用意しています。 以下は、TypeScript の同じ例であり、インターフェイスを使用したものです。

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

親関数が非同期の場合は **requestUpdate()** を呼び出すことができます`await`が、Office アプリケーションはリボンの状態を更新するタイミングを制御します。 **requestUpdate ()** メソッドが、更新の要求をキューイングします。 このメソッドは、リボンが実際に更新されたときではなく、要求をキューに入れるとすぐに promise オブジェクトを解決します。

## <a name="change-the-state-in-response-to-an-event"></a>イベントに応じて状態を変更する

リボンの状態を変更する一般的なシナリオは、ユーザーが開始したイベントがアドインのコンテキストを変更したときです。

グラフがアクティブになったときにのみボタンを有効にするシナリオを考えます。 まず、マニフェストのボタンの [Enabled](/javascript/api/manifest/enabled) 要素を `false` に設定します。 例については上記を参照してください。

次に、ハンドラーを割り当てます。 これは、ワークシート内のすべてのグラフの **onActivated** イベントと **onDeactivated** イベントにハンドラー (後の手順で作成) を割り当てる次の例のように **、Office.onReady** 関数で一般的に行われます。

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

そして、`enableChartFormat` ハンドラーを定義します。 以下は簡単な例ですが、より信頼性の高い方法でコントロールの状態を変更する場合については、後述の「[ベスト プラクティス: コントロールの状態エラーのテスト](#best-practice-test-for-control-status-errors)」を参照してください。

```javascript
function enableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

最後に、`disableChartFormat` ハンドラーを定義します。 `enableChartFormat` と同じですが、ボタン オブジェクトの **enabled** プロパティを `false` に設定する必要があります。

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効状態を同時に切り替える

**requestUpdate** メソッドは、カスタム コンテキスト タブの表示を切り替えるためにも使用されます。このコードとコード例の詳細については、「[Office アドインでカスタム コンテキスト タブを作成する](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)」を参照してください。

## <a name="best-practice-test-for-control-status-errors"></a>ベスト プラクティス: コントロールの状態エラーのテスト

状況によっては、`requestUpdate` が呼び出された後でもリボンが再描画されず、コントロールのクリック可能な状態が変更されない場合があります。 そこで、アドインのベスト プラクティスとして、コントロールの状態を追跡することが挙げられます。 アドインは、次の規則に従う必要があります。

1. `requestUpdate` が呼び出された場合はいつでも、コードがカスタム ボタンとメニュー項目の意図した状態を記録する必要があります。
2. カスタム コントロールがクリックされたら、ハンドラーの最初のコードが、ボタンがクリック可能になっているかどうかを確認する必要があります。 クリック可能でない場合は、コードがエラーの報告または記録を行い、ボタンを意図した状態に設定し直す必要があります。

次の例は、ボタンを無効にし、ボタンの状態を記録する関数を示しています。 `chartFormatButtonEnabled` は、マニフェスト内のボタンの [Enabled](/javascript/api/manifest/enabled) 要素と同じ値に初期化されるグローバルなブール変数です。

```javascript
function disableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

次の例は、ボタンのハンドラーがボタンの不正な状態をテストする方法を示しています。 `reportError` は、エラーを表示または記録する関数です。

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a>エラー処理

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 このエラーの処理方法の例を次に示します。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

```javascript
function disableChartFormat() {
    try {
        const button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        const parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        const parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        const ribbonUpdater = {tabs: [parentTab]};
        Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
