---
title: Office アドインでカスタム コンテキスト タブを作成する
description: Office アドインにカスタム コンテキスト タブを追加する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a079930bbb4523893f25604aefcff0a68f0316b
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889192"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Office アドインでカスタム コンテキスト タブを作成する

コンテキスト タブは、Office リボンの非表示のタブ コントロールであり、Office ドキュメントで指定されたイベントが発生したときにタブ行に表示されます。 たとえば、テーブルが選択されたときに Excel リボンに表示される [テーブル **デザイン** ] タブです。 Office アドインにカスタム コンテキスト タブを含め、表示を変更するイベント ハンドラーを作成して、表示または非表示のタイミングを指定します。 (ただし、カスタム コンテキスト タブはフォーカスの変更に応答しません)。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

> [!IMPORTANT]
> カスタム コンテキスト タブは現在、Excel でのみサポートされており、これらのプラットフォームとビルドでのみサポートされています。
>
> - Excel on Windows (Microsoft 365 サブスクリプションのみ): バージョン 2102 (ビルド 13801.20294) 以降。
> - Excel on Mac: バージョン 16.53.806.0 以降。
> - Excel on the web

> [!NOTE]
> カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ機能します。 要件セットとその使用方法の詳細については、「 [Office アプリケーションと API 要件の指定](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> コード内のランタイム チェックを使用して、メソッドと要件セットのサポートに関するランタイム チェックの説明に従って、ユーザーのホストとプラットフォームの組み合わせがこれらの [要件セットをサポートしているかどうかをテスト](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)できます。 (この記事でも説明されているマニフェストで要件セットを指定する手法は、現在 RibbonApi 1.2 では機能しません)。または、 [カスタム コンテキスト タブがサポートされていない場合は、代替 UI エクスペリエンスを実装](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)することもできます。

## <a name="behavior-of-custom-contextual-tabs"></a>カスタム コンテキスト タブの動作

カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みの Office コンテキスト タブのパターンに従います。 配置カスタム コンテキスト タブの基本原則を次に示します。

- カスタム コンテキスト タブが表示されると、リボンの右端に表示されます。
- 1 つ以上の組み込みコンテキスト タブとアドインの 1 つ以上のカスタム コンテキスト タブが同時に表示されている場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。
- アドインに複数のコンテキスト タブがあり、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。 (方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右ですが、右から左の言語では右から左です)。 [定義方法の詳細については、「タブに表示されるグループとコントロールを](#define-the-groups-and-controls-that-appear-on-the-tab) 定義する」を参照してください。
- 複数のアドインに特定のコンテキストで表示されるコンテキスト タブがある場合は、アドインが起動された順序で表示されます。
- カスタム コア タブとは異なり、カスタム *コンテキスト* タブは Office アプリケーションのリボンに永続的に追加されません。 これらは、アドインが実行されている Office ドキュメントにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキスト タブを含める主な手順

アドインにカスタム コンテキスト タブを含める主な手順を次に示します。

1. 共有ランタイムを使用するようにアドインを構成します。
1. タブとそのタブに表示されるグループとコントロールを定義します。
1. コンテキスト タブを Office に登録します。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

カスタム コンテキスト タブを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェストで XML で定義されるカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB で定義されます。 コードは BLOB を JavaScript オブジェクトに解析し、そのオブジェクトを [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) メソッドに渡します。 カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。 これは、アドインのインストール時に Office アプリケーション リボンに追加されるカスタム コア タブとは異なり、別のドキュメントを開いたときに表示されたままになります。 また、このメソッドは `requestCreateControls` 、アドインのセッションで 1 回だけ実行できます。 再度呼び出されると、エラーがスローされます。

> [!NOTE]
> JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML 内の [CustomTab](/javascript/api/manifest/customtab) 要素とその子孫要素の構造とほぼ平行です。

コンテキスト タブの JSON BLOB のステップ バイ ステップの例を作成します。 コンテキスト タブ JSON の完全なスキーマは [、dynamic-ribbon.schema.json にあります](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 Visual Studio Code で作業している場合は、このファイルを使用して IntelliSense を取得し、JSON を検証できます。 詳細については、「 [Visual Studio Code を使用した JSON の編集 - JSON スキーマと設定](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)」を参照してください。

1. まず、名前`actions``tabs`と . 配列 `actions` は、コンテキスト タブのコントロールによって実行できるすべての関数の仕様です。配列は `tabs` 、 *最大 20* 個までのコンテキスト タブを定義します。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. コンテキスト タブのこの簡単な例では、ボタンは 1 つしか持たないため、アクションは 1 つだけです。 配列の唯一のメンバーとして次を `actions` 追加します。 このマークアップについては、次の点に注意してください。

    - プロパティと`type`プロパティは`id`必須です。
    - の値 `type` は、"ExecuteFunction" または "ShowTaskpane" のいずれかです。
    - プロパティは`functionName`、値`type``ExecuteFunction`が . これは、FunctionFile で定義されている関数の名前です。 FunctionFile の詳細については、「 [アドイン コマンドの基本的な概念」を参照してください](add-in-commands.md)。
    - 後の手順では、このアクションをコンテキスト タブのボタンにマップします。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. 配列の唯一のメンバーとして次を `tabs` 追加します。 このマークアップについては、次の点に注意してください。

    - `id` プロパティは必須です。 アドイン内のすべてのコンテキスト タブで一意の簡潔で説明的な ID を使用します。
    - `label` プロパティは必須です。 コンテキスト タブのラベルとして機能するユーザー フレンドリな文字列です。
    - `groups` プロパティは必須です。 タブに表示されるコントロールのグループを定義します。メンバーは少なくとも 1 つ *、20 以下* である必要があります。 (カスタム コンテキスト タブで使用できるコントロールの数にも制限があり、持つグループの数も制限されます。 詳細については、次の手順を参照してください)。

    > [!NOTE]
    > タブ オブジェクトには、アドインの起動時にタブをすぐに表示するかどうかを指定する省略可能 `visible` なプロパティを持つこともできます。 コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になるため (ドキュメント内の何らかの種類のエンティティを選択するユーザーなど)、プロパティは `visible` 既定で存在しない場合に `false` 設定されます。 後のセクションでは、イベントに応答してプロパティを設定する `true` 方法について説明します。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 単純な進行中の例では、コンテキスト タブには 1 つのグループしかありません。 配列の唯一のメンバーとして次を `groups` 追加します。 このマークアップについては、次の点に注意してください。

    - すべてのプロパティが必要です。
    - このプロパティは `id` 、マニフェスト内のすべてのグループで一意である必要があります。 最大 125 文字の簡潔でわかりやすい ID を使用します。
    - グループ `label` のラベルとして機能するユーザー フレンドリな文字列です。
    - `icon`プロパティの値は、リボンと Office アプリケーション ウィンドウのサイズに応じて、グループがリボンに表示するアイコンを指定するオブジェクトの配列です。
    - `controls`プロパティの値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。 少なくとも 1 つ存在する必要があります。

    > [!IMPORTANT]
    > *タブ全体のコントロールの合計数は 20 以下にできます。* たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと 2 つのコントロールを持つ 4 番目のグループを持つことができますが、それぞれ 6 つのコントロールを持つ 4 つのグループを持つことはできません。  

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

1. すべてのグループには、32 x 32 ピクセルと 80x80 ピクセルの 2 つ以上のサイズのアイコンが必要です。 必要に応じて、サイズ 16x16 ピクセル、20x20 ピクセル、24x24 ピクセル、40x40 ピクセル、48x48 ピクセル、64x64 ピクセルのアイコンを設定することもできます。 Office は、リボンと Office アプリケーション ウィンドウのサイズに基づいて、使用するアイコンを決定します。 アイコン配列に次のオブジェクトを追加します。 (ウィンドウとリボンのサイズが *グループのコントロール* の少なくとも 1 つを表示するのに十分な大きさの場合、グループ アイコンはまったく表示されません。 たとえば、Word ウィンドウを縮小して展開するときに、Word リボンの **[スタイル** ] グループを確認します)。このマークアップについては、次の点に注意してください。

    - 両方のプロパティが必要です。
    - `size`プロパティの測定単位はピクセルです。 アイコンは常に正方形であるため、数値は高さと幅の両方です。
    - このプロパティは `sourceLocation` 、アイコンの完全な URL を指定します。

    > [!IMPORTANT]
    > 通常、開発から運用環境に移行するときにアドインのマニフェストの URL を変更する必要がある場合と同様に (ドメインを localhost から contoso.com に変更するなど)、コンテキスト タブ JSON の URL も変更する必要があります。

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

1. この簡単な継続的な例では、グループに 1 つのボタンしかありません。 配列の唯一のメンバーとして次のオブジェクトを `controls` 追加します。 このマークアップについては、次の点に注意してください。

    - を除く `enabled`すべてのプロパティが必要です。
    - `type` コントロールの種類を指定します。 値には、"Button"、"Menu"、または "MobileButton" を指定できます。
    - `id` は最大 125 文字です。
    - `actionId` は、配列で `actions` 定義されているアクションの ID である必要があります。 (このセクションの手順 1 を参照してください)。
    - `label` は、ボタンのラベルとして機能するわかりやすい文字列です。
    - `superTip` は、豊富なツール ヒントの形式を表します。 プロパティと`description`プロパティの`title`両方が必要です。
    - `icon` は、ボタンのアイコンを指定します。 グループ アイコンに関する前のコメントもここに適用されます。
    - `enabled` (省略可能) は、コンテキスト タブの起動時にボタンを有効にするかどうかを指定します。 存在しない場合の既定値は `true`.

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>requestCreateControls でコンテキスト タブを Office に登録する

コンテキスト タブは、 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) メソッドを呼び出すことによって Office に登録されます。 これは通常、メソッドに割り当てられている関数またはメソッドを使用`Office.onReady`して`Office.initialize`行われます。 これらのメソッドとアドインの初期化の詳細については、「 [Office アドインを初期化](../develop/initialize-add-in.md)する」を参照してください。 ただし、初期化後はいつでもメソッドを呼び出すことができます。

> [!IMPORTANT]
> このメソッドは `requestCreateControls` 、アドインの特定のセッションで 1 回だけ呼び出されることがあります。 再度呼び出されると、エラーがスローされます。

次に例を示します。 JSON 文字列を JavaScript 関数に渡す前に、メソッドを `JSON.parse` 使用して JavaScript オブジェクトに変換する必要があることに注意してください。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>requestUpdate でタブが表示されるコンテキストを指定する

通常、ユーザーが開始したイベントがアドイン コンテキストを変更すると、カスタム コンテキスト タブが表示されます。 (Excel ブックの既定のワークシートで) グラフがアクティブ化された場合にのみタブを表示するシナリオを考えてみましょう。

ハンドラーを割り当てることから始めます。 これは、ワークシート内`Office.onReady`のすべてのグラフのハンドラー (後の手順で作成) と`onDeactivated`イベントを`onActivated`割り当てる次の例のように、メソッドで一般的に行われます。

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

次に、ハンドラーを定義します。 次に示す簡単な例を `showDataTab`示しますが、関数のより堅牢なバージョンについては、この記事の後半の [HostRestartNeeded エラーの処理](#handle-the-hostrestartneeded-error) に関するページを参照してください。 このコードについては、以下の点に注意してください。

- Office では、リボンの状態を更新するタイミングが制御されます。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) メソッドは、更新要求をキューに入れます。 メソッドは、リボンが実際に `Promise` 更新されたときではなく、要求をキューに入れるとすぐにオブジェクトを解決します。
- メソッドの `requestUpdate` パラメーターは [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) *は JSON で指定した* ID でタブを指定し、(2) はタブの可視性を指定します。
- 同じコンテキストで表示するカスタム コンテキスト タブが複数ある場合は、配列にタブ オブジェクトを `tabs` 追加するだけです。

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

タブを非表示にするハンドラーはほぼ同じですが、プロパティ`false`を `visible` .

Office JavaScript ライブラリには、オブジェクトの構築`RibbonUpdateData` を容易にするために、いくつかのインターフェイス (種類) も用意されています。 TypeScript の関数を `showDataTab` 次に示します。これらの型を使用します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効状態を同時に切り替える

この `requestUpdate` メソッドは、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替えるためにも使用されます。詳細については、「 [アドイン コマンドの有効化と無効化](disable-add-in-commands.md)」を参照してください。 タブの表示とボタンの有効状態の両方を同時に変更するシナリオが考えられます。 これは 1 回の `requestUpdate`呼び出しで行います。 次に示す例は、コンテキスト タブが表示されるのと同時に、コア タブのボタンが有効になっている例です。

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

次の例では、有効になっているボタンは、表示されているのとまったく同じコンテキスト タブにあります。

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

カスタム コンテキスト タブのボタンから作業ウィンドウを開くには、次の`ShowTaskpane`操作を含む`type`アクションを JSON で作成します。 次に、アクションのプロパティが設定された `actionId` ボタンを `id` 定義します。 これにより、マニフェスト内の要素で指定された既定の作業ウィンドウが **\<Runtime\>** 開きます。

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

既定の作業ウィンドウではない作業ウィンドウを開くには、アクションの定義にプロパティを指定 `sourceLocation` します。 次の例では、別のボタンから 2 番目の作業ウィンドウが開かれています。

> [!IMPORTANT]
>
> - アクションに a `sourceLocation` を指定すると、作業ウィンドウで共有ランタイムは使用 *されません* 。 新しい JavaScript ランタイムで実行されます。
> - 共有ランタイムを使用できる作業ウィンドウは複数ないため、プロパティを省略`sourceLocation`できる種類`ShowTaskpane`のアクションは 1 つ以上ありません。

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

渡 `requestCreateControls` される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法でローカライズされません (マニフェスト [からのローカライズの制御](../develop/localization.md#control-localization-from-the-manifest)に関するページで説明されています)。 代わりに、ロケールごとに個別の JSON BLOB を使用して、ローカライズを実行時に行う必要があります。 [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) プロパティを`switch`テストするステートメントを使用することをお勧めします。 次に例を示します。

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

次に、次の例のように、コードは関数を呼び出して `requestCreateControls`、渡されるローカライズされた BLOB を取得します。

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>カスタム コンテキスト タブのベスト プラクティス

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>カスタム コンテキスト タブがサポートされていない場合に、代替 UI エクスペリエンスを実装する

プラットフォーム、Office アプリケーション、および Office ビルドのいくつかの組み合わせはサポート `requestCreateControls`されていません。 アドインは、これらの組み合わせのいずれかでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計する必要があります。 次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。

#### <a name="use-noncontextual-tabs-or-controls"></a>非コンテキスト タブまたはコントロールを使用する

カスタム コンテキスト タブをサポートしていないアプリケーションまたはプラットフォームでアドインが実行されているときにカスタム コンテキスト タブを実装するフォールバック エクスペリエンスをアドインで作成するように設計された manifest 要素 [、OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi) があります。

この要素を使用する最も簡単な方法は、アドイン内のカスタム コンテキスト タブのリボンカスタマイズを複製するカスタム コア タブ (つまり、 *非コンテキスト* カスタム タブ) をマニフェストに定義することです。 ただし、カスタム コア タブの重複する [グループ](/javascript/api/manifest/group)、[コントロール](/javascript/api/manifest/control)、およびメニュー **\<Item\>** 要素の最初の子要素として追加`<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`します。 その効果は次のとおりです。

- アドインがカスタム コンテキスト タブをサポートするアプリケーションとプラットフォームで実行されている場合、カスタム コア グループとコントロールはリボンに表示されません。 代わりに、アドインがメソッドを呼び出すときに、カスタム コンテキスト タブが `requestCreateControls` 作成されます。
- アドインがサポート`requestCreateControls`*されていない* アプリケーションまたはプラットフォームで実行される場合、要素はカスタム コア タブに表示されます。

次に例を示します。 "MyButton" は、カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに表示されることに注意してください。 ただし、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親グループとカスタム コア タブが表示されます。

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

親グループまたはメニューにマークが付いている `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`場合、そのグループは表示されず、カスタム コンテキスト タブがサポートされていない場合、その子マークアップはすべて無視されます。 そのため、これらの子要素のいずれかが要素を持っているか **\<OverriddenByRibbonApi\>** 、その値が何であるかは関係ありません。 この影響は、メニュー項目またはコントロールがすべてのコンテキストで表示される必要がある場合、マークを付ける必要があるだけでなく、*その先祖メニューとグループもこの方法で*`<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`マークされないようにする必要があるということです。

> [!IMPORTANT]
> グループまたはメニュー`<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`*のすべての* 子要素に . 親要素が前の段落で指定された理由で `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` マークされている場合、これは意味がありません。 さらに、親タブを **\<OverriddenByRibbonApi\>** 省略した場合 (または親に `false`設定した場合) は、カスタム コンテキスト タブがサポートされているかどうかに関係なく親が表示されますが、サポートされている場合は空になります。 そのため、カスタム コンテキスト タブがサポートされているときにすべての子要素を表示しない場合は、親 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`に .

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>指定したコンテキストで作業ウィンドウを表示または非表示にする API を使用する

代わりに **\<OverriddenByRibbonApi\>**、アドインは、カスタム コンテキスト タブのコントロールの機能を複製する UI コントロールを使用して作業ウィンドウを定義できます。次 [に、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) メソッドと [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) メソッドを使用して、コンテキスト タブがサポートされている場合にコンテキスト タブが表示されたときに作業ウィンドウを表示します。 これらのメソッドの使用方法の詳細については、「 [Office アドインの作業ウィンドウを表示または非表示にする](../develop/show-hide-add-in.md)」を参照してください。

### <a name="handle-the-hostrestartneeded-error"></a>HostRestartNeeded エラーを処理する

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 コードでは、このエラーを処理する必要があります。 方法の例を次に示します。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

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
