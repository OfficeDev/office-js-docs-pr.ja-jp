---
title: カスタム コンテキスト タブをアドインOffice作成する
description: カスタム コンテキスト タブをアドインに追加するOffice説明します。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 12286ef675a938e4abd8dd3caa90cd97586cb6d7
ms.sourcegitcommit: 6a378d2a3679757c5014808ae9da8ababbfe8b16
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/15/2021
ms.locfileid: "49870638"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Office アドインでカスタム コンテキスト タブを作成する (プレビュー)

コンテキスト タブは、指定したイベントがドキュメント内で発生した場合にタブ行に表示される Office リボン内の非表示のタブ Officeです。 たとえば、テーブルが **選択されている** ときに Excel リボンに表示される [テーブルのデザイン] タブです。 表示を変更するイベント ハンドラーを作成することで、Office アドインにカスタム コンテキスト タブを含め、いつ表示または非表示にするか指定できます。 (ただし、カスタム コンテキスト タブはフォーカスの変更には応答しない)。

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

カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのコンテキスト タブのOfficeに従います。 配置カスタム コンテキスト タブの基本的な原則を次に示します。

- カスタム コンテキスト タブが表示されている場合は、リボンの右側に表示されます。
- 1 つ以上の組み込みのコンテキスト タブと、アドインの 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。
- アドインに複数のコンテキスト タブがある場合に、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。 (方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右に、右から左の言語では右から左です)。定義 [方法の詳細については、「](#define-the-groups-and-controls-that-appear-on-the-tab) タブに表示されるグループとコントロールの定義」を参照してください。
- 特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。
- カスタム *コンテキスト* タブは、カスタムコア タブとは異なり、アプリケーションのリボンに完全Office追加されません。 アドインが実行されているOfficeドキュメントにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキスト タブを含む主な手順

アドインにカスタム コンテキスト タブを含む主な手順を次に示します。

1. 共有ランタイムを使用するアドインを構成します。
1. タブと、タブに表示されるグループとコントロールを定義します。
1. 操作に応じたタブを Office。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

カスタム コンテキスト タブを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「共有ランタイム [を使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェスト内の XML で定義されたカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB を使用して定義されます。 コードは BLOB を JavaScript オブジェクトに解析し、そのオブジェクトを [Office.ribbon.requestCreateControls メソッドに渡](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) します。 カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。 これは、アドインのインストール時に Office アプリケーション リボンに追加され、別のドキュメントが開かれたときに存在し続けるカスタム コア タブとは異なります。 また、 `requestCreateControls` このメソッドはアドインのセッションで 1 回だけ実行できます。 再度呼び出された場合は、エラーがスローされます。

> [!NOTE]
> JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は [、CustomTab](../reference/manifest/customtab.md) 要素とそのマニフェスト XML 内の子孫要素の構造と大まかに平行です。

コンテキスト タブ JSON BLOB のステップ バイ ステップの例を作成します。 (コンテキスト タブ JSON の完全なスキーマは、dynamic-ribbon.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 このリンクは、コンテキスト タブのプレビュー期間の早い段階では機能しない可能性があります。 リンクが機能しない場合は、下書きページでスキーマの最新の下書 [きdynamic-ribbon.schema.jsを見つける必要があります](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)。コードで作業しているVisual Studio、このファイルを使用して JSON のIntelliSenseを取得し、JSON を検証できます。 詳細については、「コード - JSON スキーマと [設定を使用Visual Studio JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)の編集」を参照してください。


1. まず、次の 2 つの配列プロパティを持つ JSON 文字列を作成 `actions` します `tabs` 。 配列 `actions` は、操作別タブのコントロールで実行できるすべての関数の仕様です。配列は、最大 10 までの 1 つ以上のコンテキスト タブ `tabs` *を定義します*。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. この操作別タブの単純な例にはボタンが 1 つしか含めなになのに対して、アクションは 1 つのみです。 以下を配列の唯一のメンバーとして追加 `actions` します。 このマークアップについては、次の点に注意してください。

    - プロパティ `id` と `type` プロパティは必須です。
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
    - プロパティ `controls` の値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。 1 つのグループに少なくとも *1 つ、6 以下である必要があります*。

    > [!IMPORTANT]
    > *タブ全体のコントロールの総数は 20 以下です。* たとえば、各コントロールが 6 つの 3 つのグループ、2 つのコントロールを持つ 4 つ目のグループを持つ場合、6 つのコントロールを持つ 4 つのグループを持つすることはできません。  

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
    > 開発から実稼働に移行するときに(ドメインを localhost から contoso.com に変更する場合など)、アドインのマニフェストの URL を通常は変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。

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

1. この単純な例では、グループにはボタンが 1 つのみです。 次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。 このマークアップについては、次の点に注意してください。

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>requestCreateControls で操作Officeタブを登録する

コンテキスト タブは [、Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) メソッドOffice呼び出すことによって、コンテキスト タブに登録されます。 これは通常、メソッドに割り当てられている関数またはメソッドで `Office.initialize` 行 `Office.onReady` われます。 これらのメソッドとアドインの初期化の詳細については、「アドインの初期化Office [参照してください](../develop/initialize-add-in.md)。 ただし、初期化後はメソッドをいつでも呼び出す必要があります。

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

通常、カスタム コンテキスト タブは、ユーザーが開始するイベントによってアドインのコンテキストが変更されると表示されます。 (Excel ブックの既定のワークシートにある) グラフがアクティブ化されている場合にのみ、タブを表示するシナリオを考えます。

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

次に、ハンドラーを定義します。 次に示すのは単純な例ですが、より堅牢なバージョンの関数については、この記事で後の `showDataTab` [「HostRestartNeeded](#handling-the-hostrestartneeded-error) エラーの処理」を参照してください。 このコードについては、以下の点に注意してください。

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

またOffice JavaScript ライブラリには、オブジェクトの作成を容易にするためのインターフェイス (型) `RibbonUpdateData` がいくつか用意されています。 TypeScript の `showDataTab` 関数を次に示します。これらの型を使用します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効な状態を同時に切り替える

このメソッドは、カスタム コンテキスト タブまたはカスタムコア タブのカスタム ボタンの有効または無効の状態を切り替 `requestUpdate` える場合にも使用されます。詳細については、「アドイン コマンドを [有効または無効にする」を参照してください](disable-add-in-commands.md)。 タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられます。 これは、1 回の呼び出しで行います `requestUpdate` 。 次の例では、コンテキスト タブが表示されるのと同時に、コア タブのボタンが有効になります。

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

渡される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法ではローカライズされません (マニフェストからのローカライズの制御で `requestCreateControls` [説明します](../develop/localization.md#control-localization-from-the-manifest))。 代わりに、ローカライズは、ロケールごとに異なる JSON BLOB を使用して実行時に行う必要があります。 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用してください。 例を次に示します。

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

## <a name="handling-the-hostrestartneeded-error"></a>HostRestartNeeded エラーの処理

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 このエラーの処理方法の例を次に示します。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

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
