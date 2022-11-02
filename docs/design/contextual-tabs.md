---
title: Office アドインでカスタム コンテキスト タブを作成する
description: カスタム コンテキスト タブを Office アドインに追加する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f43f6ec0a6ef3faef4c5e50d5da6d124124fe92
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810233"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Office アドインでカスタム コンテキスト タブを作成する

コンテキスト タブは、Office リボンの非表示のタブ コントロールであり、Office ドキュメントで指定したイベントが発生したときにタブ行に表示されます。 たとえば、テーブルが選択されている場合に Excel リボンに表示される [テーブル **デザイン** ] タブなどです。 Office アドインにカスタム コンテキスト タブを含め、表示/非表示を指定するには、可視性を変更するイベント ハンドラーを作成します。 (ただし、カスタム コンテキスト タブはフォーカスの変更に応答しません)。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

> [!IMPORTANT]
> カスタム コンテキスト タブは現在、Excel でのみサポートされており、これらのプラットフォームとビルドでのみサポートされています。
>
> - Excel on Windows: バージョン 2102 (ビルド 13801.20294) 以降。
> - Excel on Mac: バージョン 16.53.806.0 以降。
> - Excel on the web

> [!NOTE]
> カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ機能します。 要件セットとその使用方法の詳細については、「 [Office アプリケーションと API の要件を指定する](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> コード内のランタイム チェックを使用して、「メソッドと要件セットのサポートのランタイム チェック」で説明されているように、ユーザーのホストとプラットフォームの組み合わせでこれらの [要件セットがサポート](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)されているかどうかをテストできます。 (マニフェストで要件セットを指定する手法は、この記事でも説明されていますが、RibbonApi 1.2 では現在機能しません)。または、 [カスタム コンテキスト タブがサポートされていない場合に、代替 UI エクスペリエンスを実装](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)することもできます。

## <a name="behavior-of-custom-contextual-tabs"></a>カスタム コンテキスト タブの動作

カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みの Office コンテキスト タブのパターンに従います。 配置カスタム コンテキスト タブの基本的な原則を次に示します。

- カスタム コンテキスト タブが表示されると、リボンの右端に表示されます。
- アドインの 1 つ以上の組み込みコンテキスト タブと 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。
- アドインに複数のコンテキスト タブがあり、複数のコンテキストが表示される場合は、アドインで定義されている順序で表示されます。 (方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右、右から左の言語では右から左)。 [定義方法の詳細については、「タブに表示されるグループとコントロール](#define-the-groups-and-controls-that-appear-on-the-tab) を定義する」を参照してください。
- 特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。
- カスタム *コンテキスト* タブは、カスタム コア タブとは異なり、Office アプリケーションのリボンに永続的に追加されません。 これらは、アドインが実行されている Office ドキュメントにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキスト タブを含める主な手順

アドインにカスタム コンテキスト タブを含める主な手順を次に示します。

1. 共有ランタイムを使用するようにアドインを構成します。
1. タブとそのタブに表示されるグループとコントロールを定義します。
1. コンテキスト タブを Office に登録します。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

カスタム コンテキスト タブを追加するには、アドインで [共有ランタイム](../testing/runtimes.md#shared-runtime)を使用する必要があります。 詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェストで XML で定義されるカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB で定義されます。 コードでは、BLOB を JavaScript オブジェクトに解析し、そのオブジェクトを [Office.ribbon.requestCreateControls メソッドに](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) 渡します。 カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ存在します。 これは、アドインのインストール時に Office アプリケーション リボンに追加され、別のドキュメントを開いたときに表示されるカスタム コア タブとは異なります。 また、メソッドはアドインの `requestCreateControls` セッションで 1 回だけ実行できます。 再度呼び出されると、エラーがスローされます。

> [!NOTE]
> JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML 内の [CustomTab](/javascript/api/manifest/customtab) 要素とその子孫要素の構造とほぼ平行です。

コンテキスト タブ JSON BLOB の例を段階的に作成します。 コンテキスト タブ JSON の完全なスキーマは、 [dynamic-ribbon.schema.json にあります](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 Visual Studio Code で作業している場合は、このファイルを使用して IntelliSense を取得し、JSON を検証できます。 詳細については、「 [Visual Studio Code を使用した JSON の編集 - JSON スキーマと設定](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)」を参照してください。

1. まず、 と という名前 `actions` の 2 つの配列プロパティを持つ JSON 文字列を `tabs`作成します。 配列は `actions` 、コンテキスト タブのコントロールによって実行できるすべての関数の仕様です。配列は `tabs` 、1 つ以上のコンテキスト タブ ( *最大 20* 個) を定義します。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. コンテキスト タブのこの単純な例では、ボタンが 1 つしかないため、アクションは 1 つだけです。 配列の唯一のメンバーとして、次を `actions` 追加します。 このマークアップについては、次の点に注意してください。

    - `id`プロパティと `type` プロパティは必須です。
    - の `type` 値は、"ExecuteFunction" または "ShowTaskpane" のいずれかです。
    - プロパティは `functionName` 、 の `type` 値が の `ExecuteFunction`場合にのみ使用されます。 これは、FunctionFile で定義されている関数の名前です。 FunctionFile の詳細については、「 [アドイン コマンドの基本的な概念](add-in-commands.md)」を参照してください。
    - 後の手順では、このアクションをコンテキスト タブのボタンにマップします。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. 配列の唯一のメンバーとして、次を `tabs` 追加します。 このマークアップについては、次の点に注意してください。

    - `id` プロパティは必須です。 アドイン内のすべてのコンテキスト タブで一意の簡単でわかりやすい ID を使用します。
    - `label` プロパティは必須です。 コンテキスト タブのラベルとして機能するわかりやすい文字列です。
    - `groups` プロパティは必須です。 タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバーを持 *ち、20 以下* である必要があります。 (カスタム コンテキスト タブで使用できるコントロールの数にも制限があり、グループの数も制限されます。 詳細については、次の手順を参照してください)。

    > [!NOTE]
    > Tab オブジェクトには、アドインの起動時にタブをすぐに表示するかどうかを指定する省略可能な `visible` プロパティを指定することもできます。 コンテキスト タブは通常、ユーザー イベントによって可視性がトリガーされるまで非表示になるため (ユーザーがドキュメント内の何らかの種類のエンティティを選択する場合など) `visible` 、プロパティは既定で 存在しない場合に に `false` 設定されます。 後のセクションでは、イベントに応答して プロパティを に `true` 設定する方法について説明します。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 単純な進行中の例では、コンテキスト タブのグループは 1 つだけです。 配列の唯一のメンバーとして、次を `groups` 追加します。 このマークアップについては、次の点に注意してください。

    - すべてのプロパティが必要です。
    - プロパティは `id` 、マニフェスト内のすべてのグループ間で一意である必要があります。 最大 125 文字の簡単でわかりやすい ID を使用します。
    - `label`は、グループのラベルとして機能するわかりやすい文字列です。
    - `icon`プロパティの値は、リボンと Office アプリケーション ウィンドウのサイズに応じて、グループがリボンに表示するアイコンを指定するオブジェクトの配列です。
    - `controls`プロパティの値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。 少なくとも 1 つ必要です。

    > [!IMPORTANT]
    > *タブ全体のコントロールの合計数は、20 以下にすることができます。* たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと 2 つのコントロールを持つ 4 つ目のグループを含めることができますが、それぞれ 6 つのコントロールを持つ 4 つのグループを持つことはできません。  

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

1. すべてのグループには、32x32 px と 80x80 px の少なくとも 2 つのサイズのアイコンが必要です。 必要に応じて、サイズ 16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、64x64 px のアイコンを使用することもできます。 Office は、リボンと Office アプリケーション ウィンドウのサイズに基づいて、使用するアイコンを決定します。 アイコン配列に次のオブジェクトを追加します。 (ウィンドウとリボンのサイズが、グループ上の *コントロール* の少なくとも 1 つが表示されるのに十分な大きさの場合、グループ アイコンはまったく表示されません。 たとえば、Word ウィンドウを縮小して展開するときに、Word リボンの **[スタイル]** グループを見ます)。このマークアップについては、次の点に注意してください。

    - どちらのプロパティも必要です。
    - `size`プロパティの測定単位はピクセルです。 アイコンは常に正方形であるため、数値は高さと幅の両方です。
    - プロパティは `sourceLocation` 、アイコンへの完全な URL を指定します。

    > [!IMPORTANT]
    > 通常、開発から運用環境に移行する場合 (localhost から contoso.com へのドメインの変更など) にアドインのマニフェストの URL を変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。

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

1. 簡単な進行中の例では、グループには 1 つのボタンしかありません。 次のオブジェクトを配列の唯一の `controls` メンバーとして追加します。 このマークアップについては、次の点に注意してください。

    - を除く `enabled`すべてのプロパティが必要です。
    - `type` はコントロールの種類を指定します。 値には、"Button"、"Menu"、または "MobileButton" を指定できます。
    - `id` は最大 125 文字です。
    - `actionId` は、配列で `actions` 定義されているアクションの ID である必要があります。 (このセクションの手順 1 を参照してください)。
    - `label` は、ボタンのラベルとして機能するわかりやすい文字列です。
    - `superTip` は、豊富なツール ヒントを表します。 プロパティと `description` プロパティの`title`両方が必要です。
    - `icon` ボタンのアイコンを指定します。 グループ アイコンに関する以前の説明もここに適用されます。
    - `enabled` (省略可能) コンテキスト タブが起動したときにボタンを有効にするかどうかを指定します。 存在しない場合の既定値は です `true`。

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>requestCreateControls を使用してコンテキスト タブを Office に登録する

コンテキスト タブは、 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) メソッドを呼び出すことによって Office に登録されます。 これは通常、関数に割り当てられている関数または 関数を使用して `Office.initialize` 実行されます `Office.onReady` 。 これらの関数とアドインの初期化の詳細については、「 [Office アドインの初期化](../develop/initialize-add-in.md)」を参照してください。 ただし、初期化後はいつでも メソッドを呼び出すことができます。

> [!IMPORTANT]
> メソッドは `requestCreateControls` 、アドインの特定のセッションで 1 回だけ呼び出されることがあります。 再度呼び出されると、エラーがスローされます。

次に例を示します。 JSON 文字列を JavaScript 関数に渡すには、 メソッドを使用 `JSON.parse` して JavaScript オブジェクトに変換する必要があることに注意してください。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>タブが requestUpdate で表示されるコンテキストを指定する

通常、ユーザーが開始したイベントがアドイン コンテキストを変更すると、カスタム コンテキスト タブが表示されます。 (Excel ブックの既定のワークシート上の) グラフがアクティブ化されている場合にのみ、タブを表示するシナリオを検討してください。

まず、ハンドラーを割り当てます。 これは一般的に、次の例のように関数で `Office.onReady` 実行されます。これは、ハンドラー (後の手順で作成) を `onActivated` ワークシート内のすべてのグラフの イベントと `onDeactivated` イベントに割り当てます。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

次に、ハンドラーを定義します。 の簡単な例を次に示しますが、より堅牢なバージョンの `showDataTab`関数については、この記事の後半の [「HostRestartNeededed エラーの処理](#handle-the-hostrestartneeded-error) 」を参照してください。 このコードについては、以下の点に注意してください。

- Office では、リボンの状態を更新するタイミングが制御されます。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) メソッドは、更新要求をキューに入れます。 メソッドは、リボンが実際に `Promise` 更新されたときではなく、要求をキューに入れるとすぐにオブジェクトを解決します。
- メソッドの `requestUpdate` パラメーターは [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) は *JSON で指定されたとおりに ID でタブを指定* し、(2) はタブの可視性を指定します。
- 同じコンテキストで表示する必要があるカスタム コンテキスト タブが複数ある場合は、配列に追加のタブ オブジェクトを `tabs` 追加するだけです。

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

タブを非表示にするハンドラーは、プロパティを に戻`visible``false`す点を除いて、ほぼ同じです。

また、Office JavaScript ライブラリには、オブジェクトの構築`RibbonUpdateData` を容易にするために、いくつかのインターフェイス (型) も用意されています。 TypeScript の関数を `showDataTab` 次に示します。これらの型を使用します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効な状態を同時に切り替える

メソッドは `requestUpdate` 、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替えるためにも使用されます。詳細については、「 [アドイン コマンドを有効または無効にする](disable-add-in-commands.md)」を参照してください。 タブの表示とボタンの有効な状態の両方を同時に変更するシナリオがある場合があります。 これは、 の 1 回の `requestUpdate`呼び出しで行います。 次に、コンテキスト タブと同時にコア タブのボタンを有効にする例を示します。

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

次の例では、有効になっているボタンは、表示されるコンテキスト タブとまったく同じです。

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

## <a name="open-a-task-pane-from-contextual-tabs"></a>コンテキスト タブから作業ウィンドウを開く

カスタム コンテキスト タブのボタンから作業ウィンドウを開くには、 の を使用して JSON にアクションを`type``ShowTaskpane`作成します。 次に、 プロパティが アクションの の `actionId` `id` に設定されたボタンを定義します。 これにより、マニフェスト内の 要素によって指定された既定の作業ウィンドウが **\<Runtime\>** 開きます。

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

既定の作業ウィンドウではない作業ウィンドウを開くには、アクションの定義で プロパティを指定 `sourceLocation` します。 次の例では、別のボタンから 2 つ目の作業ウィンドウが開きます。

> [!IMPORTANT]
>
> - アクションに `sourceLocation` が指定されている場合、作業ウィンドウは共有ランタイムを使用 *しません* 。 新しい個別のランタイムで実行されます。
> - 共有ランタイムを使用できる作業ウィンドウは複数ないため、1 つ以上の種類 `ShowTaskpane` のアクションで プロパティを `sourceLocation` 省略することはできません。

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a>JSON テキストをローカライズする

渡される `requestCreateControls` JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法でローカライズされません (マニフェスト [からのローカライズの制御](../develop/localization.md#control-localization-from-the-manifest)に関するページで説明されています)。 代わりに、ロケールごとに個別の JSON BLOB を使用して実行時にローカライズを行う必要があります。 [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) プロパティを`switch`テストするステートメントを使用することをお勧めします。 次に例を示します。

```javascript
function GetContextualTabsJsonSupportedLocale () {
    const displayLanguage = Office.context.displayLanguage;

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

次に、次の例のように、コードによって 関数が呼び出され、 に `requestCreateControls`渡されるローカライズされた BLOB が取得されます。

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>カスタム コンテキスト タブのベスト プラクティス

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>カスタム コンテキスト タブがサポートされていない場合に代替 UI エクスペリエンスを実装する

プラットフォーム、Office アプリケーション、および Office ビルドの組み合わせによっては、 がサポート `requestCreateControls`されていません。 アドインは、これらの組み合わせのいずれかでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計する必要があります。 以降のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。

#### <a name="use-noncontextual-tabs-or-controls"></a>非コンテキスト タブまたはコントロールを使用する

マニフェスト要素 [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi) は、カスタム コンテキスト タブをサポートしていないアプリケーションまたはプラットフォームでアドインが実行されているときにカスタム コンテキスト タブを実装するフォールバック エクスペリエンスをアドインに作成するように設計されています。

この要素を使用する最も簡単な方法は、アドイン内のカスタム コンテキスト タブのリボンのカスタマイズを複製するカスタム コア タブ (つまり、 *コンテキストに依存しない* カスタム タブ) をマニフェストに定義することです。 ただし、カスタム コア タブの重複する [グループ](/javascript/api/manifest/group)、[コントロール](/javascript/api/manifest/control)、メニュー **\<Item\>** 要素の最初の子要素としてを追加`<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`します。 その効果は次のとおりです。

- カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインが実行されている場合、カスタム コア グループとコントロールはリボンに表示されません。 代わりに、アドインが メソッドを呼び出すと、カスタム コンテキスト タブが `requestCreateControls` 作成されます。
- をサポート`requestCreateControls`*していない* アプリケーションまたはプラットフォームでアドインが実行されている場合、要素はカスタム コア タブに表示されます。

次に例を示します。 カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに "MyButton" が表示されることに注意してください。 ただし、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親グループとカスタム コア タブが表示されます。

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
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

その他の例については、「 [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi)」を参照してください。

親グループまたはメニューが で `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`マークされている場合、そのグループは表示されず、カスタム コンテキスト タブがサポートされていない場合、その子マークアップはすべて無視されます。 そのため、これらの子要素のいずれかが要素を持っているか、 **\<OverriddenByRibbonApi\>** その値が何であるかは関係ありません。 これは、メニュー項目またはコントロールをすべてのコンテキストで表示する必要がある場合は、 でマーク `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`する必要があるだけでなく、 *先祖のメニューとグループもこのようにマークする必要がない* ということです。

> [!IMPORTANT]
> グループまたはメニュー *のすべての* 子要素を で `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`マークしないでください。 前の段落で指定した理由で親要素が で `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` マークされている場合、これは無意味です。 さらに、 を親に残 **\<OverriddenByRibbonApi\>** すか に `false`設定すると、カスタム コンテキスト タブがサポートされているかどうかに関係なく親が表示されますが、サポートされている場合は空になります。 そのため、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親を で `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`マークします。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>指定したコンテキストで作業ウィンドウを表示または非表示にする API を使用する

の **\<OverriddenByRibbonApi\>** 代わりに、アドインは、カスタム コンテキスト タブでコントロールの機能を複製する UI コントロールを含む作業ウィンドウを定義できます。次に、 [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) メソッドと [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) メソッドを使用して、コンテキスト タブがサポートされている場合に作業ウィンドウを表示します。 これらの方法の使用方法の詳細については、「 [Office アドインの作業ウィンドウを表示または非表示にする](../develop/show-hide-add-in.md)」を参照してください。

### <a name="handle-the-hostrestartneeded-error"></a>HostRestartNeeded エラーを処理する

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 コードでこのエラーを処理する必要があります。 方法の例を次に示します。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
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

## <a name="resources"></a>リソース

- [コード サンプル: リボンにカスタム コンテキスト タブを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- コンテキスト タブのコミュニティ デモのサンプル

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
