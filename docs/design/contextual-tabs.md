---
title: カスタム コンテキスト タブをアドインOffice作成する
description: カスタム コンテキスト タブをアドインに追加するOffice説明します。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 0badd779f22edc9b4659908409764bea1cde44f5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237722"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Office アドインでカスタム コンテキスト タブを作成する (プレビュー)

操作依存タブは、指定したイベントがドキュメントで発生した場合にタブ行に表示される Office リボンの非表示のタブ コントロールOfficeします。 たとえば、テーブルが **選択されている** ときに Excel リボンに表示される [テーブルのデザイン] タブです。 可視性を変更するイベント ハンドラーを作成することで、Office アドインにカスタム コンテキスト タブを含め、いつ表示または非表示にするか指定できます。 (ただし、カスタム コンテキスト タブはフォーカスの変更には応答しない)。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

> [!IMPORTANT]
> カスタム コンテキスト タブはプレビュー中です。 開発環境またはテスト環境で実験してください。ただし、実稼働アドインには追加しません。
>
> カスタム コンテキスト タブは現在 Excel でのみサポートされ、次のプラットフォームとビルドでのみサポートされています。
>
> - Excel on Windows (永続的なライセンスではなく、Microsoft 365 のみ): バージョン 2011 (ビルド 13426.20274)。 Microsoft 365 サブスクリプションは、以前は 「月次チャネル (対象指定)」または「Insider Slow」と呼ばばされている現在のチャネル [(プレビュー)](https://insider.office.com/join/windows) に登録する必要がある場合があります。

> [!NOTE]
> カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ動作します。 要件セットとそれらを使用する方法の詳細については、「アプリケーションと API の要件Office指定する」 [を参照してください](../develop/specify-office-hosts-and-api-requirements.md)。
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>カスタム コンテキスト タブの動作

カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのコンテキスト タブのパターンOfficeに従います。 配置カスタム コンテキスト タブの基本的な原則を次に示します。

- カスタム コンテキスト タブが表示されている場合は、リボンの右側に表示されます。
- 1 つ以上の組み込みのコンテキスト タブと、アドインの 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。
- アドインに複数のコンテキスト タブがある場合に、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。 (方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右に、右から左の言語では右から左です)。定義 [方法の詳細については、「](#define-the-groups-and-controls-that-appear-on-the-tab) タブに表示されるグループとコントロールの定義」を参照してください。
- 特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。
- カスタム *コンテキスト* タブは、カスタム コア タブとは異なり、アプリケーションのリボンに完全Office追加されません。 アドインが実行されているOfficeドキュメントにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキスト タブを含む主な手順

アドインにカスタム コンテキスト タブを含む主な手順を次に示します。

1. 共有ランタイムを使用するアドインを構成します。
1. タブと、タブに表示されるグループとコントロールを定義します。
1. コンテキスト タブをユーザー設定にOffice。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

カスタム コンテキスト タブを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「共有ランタイム [を使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェスト内の XML で定義されたカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB を使用して定義されます。 コードは BLOB を JavaScript オブジェクトに解析し、そのオブジェクトを [Office.ribbon.requestCreateControls メソッドに渡](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) します。 カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。 これは、アドインのインストール時に Office アプリケーション リボンに追加され、別のドキュメントが開かれたときに存在し続けるカスタム コア タブとは異なります。 また、 `requestCreateControls` このメソッドはアドインのセッションで 1 回だけ実行できます。 再度呼び出された場合は、エラーがスローされます。

> [!NOTE]
> JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は [、CustomTab](../reference/manifest/customtab.md) 要素とそのマニフェスト XML 内の子孫要素の構造と大まかに平行です。

コンテキスト タブ JSON BLOB のステップ バイ ステップで例を作成します。 (コンテキスト タブ JSON の完全なスキーマは、dynamic-ribbon.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 このリンクは、コンテキスト タブのプレビュー期間中に機能しない可能性があります。 リンクが機能しない場合は、下書きページでスキーマの最新の下書 [きdynamic-ribbon.schema.jsを見つける必要があります](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json)。コードで作業している場合Visual Studioこのファイルを使用して、JSON IntelliSenseを取得し、検証できます。 詳細については、「コード - JSON スキーマと [設定を使用Visual Studio JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)の編集」を参照してください。


1. まず、次の 2 つの配列プロパティを持つ JSON 文字列を作成 `actions` します `tabs` 。 配列 `actions` は、操作別タブのコントロールで実行できるすべての関数の仕様です。配列 `tabs` は、最大 *20* までの 1 つ以上のコンテキスト タブを定義します。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. この操作別タブの単純な例にはボタンが 1 つしか含めなく、したがってアクションは 1 つのみです。 以下を配列の唯一のメンバーとして追加 `actions` します。 このマークアップについては、次の点に注意してください。

    - プロパティ `id` `type` とプロパティは必須です。
    - 値には `type` 、"ExecuteFunction" または "ShowTaskpane" を指定できます。
    - プロパティ `functionName` は、値が次の場合にのみ使用 `type` されます `ExecuteFunction` 。 FunctionFile で定義されている関数の名前です。 FunctionFile の詳細については、「アドイン コマンドの基本 [概念」を参照してください](add-in-commands.md)。
    - 後の手順では、このアクションをコンテキスト タブのボタンにマップします。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. 以下を配列の唯一のメンバーとして追加 `tabs` します。 このマークアップについては、次の点に注意してください。

    - `id` プロパティは必須です。 アドイン内のすべてのコンテキスト タブの中で一意である簡潔でわかりやすい ID を使用します。
    - `label` プロパティは必須です。 コンテキスト タブのラベルとして使用すると、ユーザーに分け親しまれる文字列になります。
    - `groups` プロパティは必須です。 タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバーと *20 以下である必要があります*。 (カスタム コンテキスト タブに設定できるコントロールの数にも制限があります。また、持っているグループの数も制限されます。 詳細については、次の手順を参照してください)。

    > [!NOTE]
    > タブ オブジェクトには、アドインの起動直後にタブを表示するかどうかを指定するオプションのプロパティ `visible` を指定することもできます。 コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になります (ユーザーがドキュメント内の何らかの種類のエンティティを選択した場合など)、プロパティは既定で存在しない場合に設定されます `visible` `false` 。 後のセクションでは、イベントに応答してプロパティを `true` 設定する方法について説明します。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 単純な例では、コンテキスト タブには 1 つのグループのみがあります。 以下を配列の唯一のメンバーとして追加 `groups` します。 このマークアップについては、次の点に注意してください。

    - すべてのプロパティが必要です。
    - この `id` プロパティは、タブ内のすべてのグループ間で一意である必要があります。簡潔でわかりやすい ID を使用します。
    - グループ `label` のラベルとして使用する、ユーザー に分かしい文字列です。
    - プロパティの値は、リボンのサイズとアプリケーション ウィンドウのサイズに応じてリボンに表示されるアイコンを指定するOffice `icon` 配列です。
    - プロパティ `controls` の値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。 少なくとも 1 つが必要です。

    > [!IMPORTANT]
    > *タブ全体のコントロールの総数は 20 以下です。* たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと 2 つのコントロールを持つ 4 つ目のグループを持つ場合がありますが、4 つのグループにそれぞれ 6 つのコントロールを持つすることはできません。  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. すべてのグループには、32x32 px と 80x80 px の 2 つ以上のサイズのアイコンが必要です。 必要に応じて、16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、64x64 px のアイコンを設定できます。 Office、リボンとアプリケーション ウィンドウのサイズに基づいて使用するアイコンOffice決定します。 アイコン配列に次のオブジェクトを追加します。 (ウィンドウとリボンのサイズが、グループ上のコントロールの少なくとも1 つが表示されるのに十分な大きさの場合、グループ アイコンは表示されません。 たとえば、Word **ウィンドウを縮小** して展開する場合は、Word リボンの [スタイル] グループを参照してください)。このマークアップについては、次の点に注意してください。

    - 両方のプロパティが必要です。
    - プロパティ `size` の単位はピクセルです。 アイコンは常に正方形なので、数値は高さと幅の両方です。
    - この `sourceLocation` プロパティは、アイコンの完全な URL を指定します。

    > [!IMPORTANT]
    > 開発から実稼働に移行する場合 (ドメインを localhost から contoso.com に変更する場合など) アドインのマニフェストの URL を通常は変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. この単純な例では、グループにボタンが 1 つしか表示されます。 次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。 このマークアップについては、次の点に注意してください。

    - ただし、すべてのプロパティ `enabled` は必須です。
    - `type` コントロールの種類を指定します。 値には、"Button"、"Menu"、または "MobileButton" を指定できます。
    - `id` 最大 125 文字まで入力できます。 
    - `actionId` は、配列で定義されているアクションの ID である必要 `actions` があります。 (このセクションの手順 1 を参照してください)。
    - `label` は、ボタンのラベルとして使用する、ユーザー に使い分け可能な文字列です。
    - `superTip` は、豊富な形式のツール ヒントを表します。 プロパティと `title` プロパティ `description` の両方が必要です。
    - `icon` ボタンのアイコンを指定します。 グループ アイコンに関する前の注釈もここに適用されます。
    - `enabled` (オプション) コンテキスト タブが表示される際にボタンを有効にするかどうかを指定します。 存在しない場合の既定値は次の値です `true` 。 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
JSON BLOB の完全な例を次に示します。

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>requestCreateControls を使用してOfficeタブを登録する

コンテキスト タブは [、Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) メソッドをOfficeして、コンテキスト タブに登録されます。 これは通常、メソッドに割り当てられている関数またはメソッドで `Office.initialize` 行 `Office.onReady` われます。 これらのメソッドとアドインの初期化の詳細については、「アドインの初期化Office [参照してください](../develop/initialize-add-in.md)。 ただし、初期化後はメソッドを呼び出す必要があります。

> [!IMPORTANT]
> この `requestCreateControls` メソッドは、アドインの特定のセッションで 1 回だけ呼び出されます。 再度呼び出された場合は、エラーがスローされます。

次に例を示します。 JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があります。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>requestUpdate でタブが表示されるコンテキストを指定する

通常、カスタム コンテキスト タブは、ユーザーが開始するイベントによってアドインのコンテキストが変更されると表示されます。 (Excel ブックの既定のワークシートにある) グラフがアクティブ化されている場合にのみ、タブが表示されるシナリオを考えます。

まず、ハンドラーを割り当てる必要があります。 これは通常、次の例のようにメソッドで行われます。この例では、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフのイベントに割り当 `Office.onReady` `onActivated` `onDeactivated` てる必要があります。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

次に、ハンドラーを定義します。 次に示すのは単純な例ですが、より堅牢なバージョンの関数については、この記事で後の `showDataTab` [「HostRestartNeeded](#handle-the-hostrestartneeded-error) エラーの処理」を参照してください。 このコードについては、以下の点に注意してください。

- Office では、リボンの状態を更新するタイミングが制御されます。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)メソッドは、更新要求をキューに入れられます。 このメソッドは、リボンが実際に更新されるのではなく、要求をキューに入れ次第、オブジェクト `Promise` を解決します。
- メソッドのパラメーターは `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) *JSON* で指定されているとおりに ID でタブを指定し、(2) タブの可視性を指定します。
- 同じコンテキストで表示するカスタム コンテキスト タブが複数ある場合は、単純にタブ オブジェクトを配列に追加 `tabs` します。

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

タブを非表示にするハンドラーは、プロパティを設定し戻す以外は、ほぼ `visible` 同じです `false` 。

またOffice JavaScript ライブラリには、オブジェクトの作成を容易にするためのインターフェイス (型) `RibbonUpdateData` がいくつか用意されています。 TypeScript の `showDataTab` 関数を次に示します。この関数は、これらの型を利用します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効な状態を同時に切り替える

このメソッドは、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替 `requestUpdate` える場合にも使用されます。詳細については、「アドイン コマンドを [有効または無効にする」を参照してください](disable-add-in-commands.md)。 タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられます。 これは、1 回の呼び出しで行います `requestUpdate` 。 次の例では、コンテキスト タブが表示されるのと同時に、コア タブのボタンが有効になります。

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
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
            ]}
        ]
    });
}
```

次の例では、有効になっているボタンは、表示されているのと同じコンテキスト タブにあります。

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
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

## <a name="localizing-the-json-blob"></a>JSON BLOB のローカライズ

渡される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法ではローカライズされません (マニフェストからのコントロールのローカライズで `requestCreateControls` [説明します](../develop/localization.md#control-localization-from-the-manifest))。 代わりに、ローカライズは、ロケールごとに異なる JSON BLOB を使用して実行時に行う必要があります。 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用してください。 例を次に示します。

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

次に、次の例のように、コードで関数を呼び出して、渡されるローカライズされた BLOB `requestCreateControls` を取得します。

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>カスタム コンテキスト タブのベスト プラクティス

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する

プラットフォーム、アプリケーション、Office、およびOfficeの組み合わせはサポートされていません `requestCreateControls` 。 アドインは、これらの組み合わせの 1 つでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計する必要があります。 次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。

#### <a name="use-noncontextual-tabs-or-controls"></a>コンテキスト以外のタブまたはコントロールを使用する

マニフェスト要素 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)は、カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するアドインでフォールバック エクスペリエンスを作成するように設計されています。 

この要素を使用する最も簡単な方法は、アドインのカスタム コンテキスト タブのリボンカスタマイズを複製する 1 つ以上のカスタム コア タブ (つまり、非コンテキスト カスタム タブ) をマニフェストで定義する方法です。 ただし `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` [、CustomTab](../reference/manifest/customtab.md)の最初の子要素として追加します。 その効果は次のとおりです。

- カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインを実行する場合、カスタムコア タブはリボンに表示されません。 代わりに、アドインがメソッドを呼び出す際にカスタム コンテキスト タブが作成 `requestCreateControls` されます。
- アドインがサポートしていないアプリケーションまたはプラットフォームで実行される場合、カスタム コア `requestCreateControls` タブがリボンに表示されます。

この簡単な戦略の例を次に示します。

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
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

この簡単な戦略では、カスタム コンテキスト タブと子グループとコントロールをミラー化するカスタム コア タブを使用しますが、より複雑な戦略を使用できます。 要素は、Group 要素と Control 要素 (ボタンの種類とメニューの種類の両方) とメニュー要素に (最初の) 子要素として `<OverriddenByRibbonApi>` 追加[](../reference/manifest/control.md#button-control)[](../reference/manifest/group.md)[](../reference/manifest/control.md)[](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` することもできます。 この事実により、コンテキスト タブに表示されるグループとコントロールを、さまざまなカスタム コア タブのさまざまなグループ、ボタン、メニューに分散できます。 次に例を示します。 "MyButton" は、カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに表示されます。 ただし、カスタム コンテキスト タブがサポートされるかどうかに関係なく、親グループとカスタムコア タブが表示されます。

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
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

その他の例については [、「OverriddenByRibbonApi」を参照してください](../reference/manifest/overriddenbyribbonapi.md)。

親タブ、グループ、またはメニューにマークが付いている場合、そのタブは表示されません。カスタム コンテキスト タブがサポートされていない場合、そのすべての子マークアップは無視されます `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 そのため、これらの子要素の中に要素がある場合や、その値 `<OverriddenByRibbonApi>` が何かは関係ありません。 これは、メニュー項目、コントロール、またはグループをすべてのコンテキストで表示する必要がある場合、メニュー項目、コントロール、またはグループがマークされていないだけでなく、その先祖のメニュー、グループ、およびタブもこの方法でマークされなければならないという意味です `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 

> [!IMPORTANT]
> タブ、グループ *、または* メニューのすべての子要素にマークを付けはしない `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 前の段落で説明した理由で親要素にマークが付いている場合、 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` これは無意味です。 さらに、親のタブを指定しない (または親に設定した) 場合は、カスタム コンテキスト タブがサポートされているかどうかに関係なく親が表示されますが、サポートされている場合は空になります。 `<OverriddenByRibbonApi>` `false` したがって、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親と親のみをマークします `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>指定したコンテキストで作業ウィンドウを表示または非表示にする API を使用する

代わりに、アドインは、カスタム コンテキスト タブのコントロールの機能を複製する UI コントロールを含む作業ウィンドウ `<OverriddenByRibbonApi>` を定義できます。 [次に、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) メソッドと [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) メソッドを使用して、操作別タブがサポートされている場合にのみ作業ウィンドウを表示します。 これらのメソッドの使い方の詳細については、アドインの作業ウィンドウを表示または非表示にするOffice [参照してください](../develop/show-hide-add-in.md)。

### <a name="handle-the-hostrestartneeded-error"></a>HostRestartNeeded エラーの処理

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 コードでこのエラーを処理する必要があります。 その方法の例を次に示します。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
