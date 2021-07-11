---
title: アドイン コマンドを有効または無効にする
description: Office Web アドインのカスタム リボン ボタンとメニュー項目の有効または無効の状態を変更する方法について説明します。
ms.date: 04/30/2021
localization_priority: Normal
ms.openlocfilehash: 2a2816990a7f21a4238a9f8332537bf904fa4cb2
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349225"
---
# <a name="enable-and-disable-add-in-commands"></a>アドイン コマンドを有効または無効にする

アドインの一部の機能を特定のコンテキストでのみ使用可能にする必要がある場合、カスタム アドイン コマンドをプログラムで有効または無効にすることができます。 たとえば、表の見出しを変更する関数は、カーソルが表の中にある場合にのみ有効にする必要があります。

また、クライアント アプリケーションを開く際にコマンドを有効または無効Office指定することもできます。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Officeとプラットフォームのサポートのみ

この記事で説明する API は、すべてのプラットフォームExcelおよびすべてのプラットフォームでのみPowerPoint on the web。

### <a name="test-for-platform-support-with-requirement-sets"></a>要件セットを使用したプラットフォーム サポートのテスト

要件セットは、API メンバーの名前付きグループです。 Officeアドインは、マニフェストで指定された要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションとプラットフォームの組み合わせがアドインに必要な API をサポートするかどうかを判断します。 詳細については、「バージョンと要件[Officeを参照してください](../develop/office-versions-and-requirement-sets.md)。

有効/無効 API は [、RibbonApi 1.1 要件セットに](../reference/requirement-sets/ribbon-api-requirement-sets.md) 属します。

> [!NOTE]
> **RibbonApi 1.1** 要件セットはマニフェストでまだサポートされていないので、マニフェストのセクションで指定 `<Requirements>` することはできません。 サポートをテストするには、コードを呼び出す必要があります `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` 。 その呼 *び出しが返* された場合にのみ、 `true` コードは有効/無効 API を呼び出す可能性があります。 戻り値の `isSetSupported` 呼び出しの場合は、すべてのカスタム アドイン `false` コマンドがすべての時間有効になります。 **RibbonApi 1.1** 要件セットがサポートされていない場合の動作を考慮に入れて、実稼働アドインとアプリ内命令を設計する必要があります。 使用の詳細と例については、「アプリケーションと API Officeを指定する、特に JavaScript コードでランタイム チェックを使用する」 `isSetSupported` [を参照してください](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)。 [](../develop/specify-office-hosts-and-api-requirements.md) (この記事 [のマニフェストで Requirements 要素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) を設定するセクションは、リボン 1.1 には適用されません)。

## <a name="shared-runtime-required"></a>共有ランタイムが必要

この記事で説明されている API とマニフェストのマークアップでは、アドインのマニフェストで共有ランタイムを使用するよう指定されている必要があります。 これを行うには、次の手順を実行します。

1. マニフェストの [Runtimes](../reference/manifest/runtimes.md) 要素で、子要素の `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />` を追加します。 (マニフェストに `<Runtimes>` 要素がまだない場合は、`VersionOverrides` セクションの `<Host>` 要素の下に最初の子要素として作成します。)
2. マニフェストの [Resources.Urls](../reference/manifest/resources.md) セクションで、子要素の `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />` を追加します。ここでは、`{MyDomain}` はアドインのドメインで、`{path-to-start-page}` はアドインの開始ページのパスになります (例: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`)。
3. アドインに作業ウィンドウ、関数ファイル、または Excel カスタム関数が含まれているかどうかに応じて、次の 3 つの手順の 1 つ以上を実行する必要があります。

    - アドインに作業ウィンドウが含まれている場合は、[Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) 要素の `resid` 属性を、手順 1 で `<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。 そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。
    - アドインに Excel カスタム関数が含まれている場合は、[Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) 要素の `resid` 属性を、手順 1 で`<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。 そうすると要素は `<SourceLocation resid="Contoso.SharedRuntime.Url"/>` のようになります。
    - アドインに関数ファイルが含まれている場合は、[FunctionFile](../reference/manifest/functionfile.md) 要素の `resid` 属性を、手順 1 で `<Runtime>` 要素の `resid` に使用したのとまったく同じ文字列に設定します。たとえば、`Contoso.SharedRuntime.Url` のようにします。 そうすると要素は `<FunctionFile resid="Contoso.SharedRuntime.Url"/>` のようになります。

## <a name="set-the-default-state-to-disabled"></a>既定の状態を無効に設定する

既定では、Office アプリケーションの起動時にすべてのアドイン コマンドが有効になります。 Office アプリケーションの起動時にカスタム ボタンまたはメニュー項目を無効にするには、マニフェストで指定します。 コントロールの宣言の [Action](../reference/manifest/action.md) 要素の *直下* (内部ではない) に、[Enabled](../reference/manifest/enabled.md) 要素 (値は `false`) を追加するだけで無効にすることができます。 以下に、基本的な構造を示します。

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
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>プログラムで状態を変更する

アドイン コマンドの有効な状態を変更するには、以下の手順が重要になります。

1. (1) コマンドとその親グループとタブをマニフェストで宣言された名前の ID で指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトを作成します。および (2) は、コマンドの有効または無効の状態を指定します。
2. **RibbonUpdaterData** オブジェクトを [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) メソッドに渡します。

次に簡単な例を示します。 マニフェストから "MyButton"、"OfficeAddinTab1"、および "CustomGroup111" がコピーされます。

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

親関数 `await` が非同期である **場合は requestUpdate()** の呼び出しを実行できますが、リボンの状態を更新するときに、Office アプリケーションが制御する点に注意してください。 **requestUpdate ()** メソッドが、更新の要求をキューイングします。 このメソッドは、リボンが実際に更新された場合ではなく、要求をキューに入れ次第、promise オブジェクトを解決します。

## <a name="change-the-state-in-response-to-an-event"></a>イベントに応じて状態を変更する

リボンの状態を変更する一般的なシナリオは、ユーザーが開始したイベントがアドインのコンテキストを変更したときです。

グラフがアクティブになったときにのみボタンを有効にするシナリオを考えます。 まず、マニフェストのボタンの [Enabled](../reference/manifest/enabled.md) 要素を `false` に設定します。 例については上記を参照してください。

次に、ハンドラーを割り当てます。 これは通常、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフの **onActivated** および **onDeactivated** イベントに割り当てる以下の例のように、**Office.onReady** メソッドで行います。

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
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
    var button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

最後に、`disableChartFormat` ハンドラーを定義します。 `enableChartFormat` と同じですが、ボタン オブジェクトの **enabled** プロパティを `false` に設定する必要があります。

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効な状態を同時に切り替える

requestUpdate **メソッド** は、カスタム コンテキスト タブの表示を切り替える場合にも使用されます。このコードとコード例の詳細については、「カスタム コンテキスト タブを作成する」を参照 [Officeアドインを参照してください](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time)。

## <a name="best-practice-test-for-control-status-errors"></a>ベスト プラクティス: コントロールの状態エラーのテスト

状況によっては、`requestUpdate` が呼び出された後でもリボンが再描画されず、コントロールのクリック可能な状態が変更されない場合があります。 そこで、アドインのベスト プラクティスとして、コントロールの状態を追跡することが挙げられます。 アドインは次のルールに準拠する必要があります。

1. `requestUpdate` が呼び出された場合はいつでも、コードがカスタム ボタンとメニュー項目の意図した状態を記録する必要があります。
2. カスタム コントロールがクリックされたら、ハンドラーの最初のコードが、ボタンがクリック可能になっているかどうかを確認する必要があります。 クリック可能でない場合は、コードがエラーの報告または記録を行い、ボタンを意図した状態に設定し直す必要があります。

次の例は、ボタンを無効にし、ボタンの状態を記録する関数を示しています。 `chartFormatButtonEnabled` は、マニフェスト内のボタンの [Enabled](../reference/manifest/enabled.md) 要素と同じ値に初期化されるグローバルなブール変数です。

```javascript
function disableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
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
        var button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        var parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        var parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        var ribbonUpdater = {tabs: [parentTab]};
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
