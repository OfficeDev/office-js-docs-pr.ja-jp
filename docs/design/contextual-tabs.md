---
title: カスタム コンテキスト タブを Officeアドインで作成する
description: カスタム コンテキスト タブをアドインに追加するOffice説明します。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7a2c6c93c009b42e1017bd52272ff0cb8a60085e
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222137"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>カスタム コンテキスト タブを Officeアドインで作成する

コンテキスト タブは、指定したイベントがドキュメント内で発生した場合にタブ行に表示Officeリボンの非表示のタブ コントロールOfficeです。 たとえば、テーブル **が選択されている** ときにリボンのExcel[テーブルのデザイン] タブが表示されます。 カスタム コンテキスト タブは、Officeアドインに含め、表示設定を変更するイベント ハンドラーを作成して、表示または非表示の状態を指定します。 (ただし、カスタム コンテキスト タブはフォーカスの変更に応答しない)。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

[!INCLUDE [Animation of contextual tabs and enabling buttons](../includes/animation-contextual-tabs-enable-button.md)]

> [!IMPORTANT]
> カスタム コンテキスト タブは現在、これらのプラットフォームExcelビルドでのみサポートされています。
>
> - Excel (Windows サブスクリプションMicrosoft 365): バージョン 2102 (ビルド 13801.20294) 以降。
> - Excel on the web

> [!NOTE]
> カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ機能します。 要件セットとそれらを使用する方法の詳細については、「アプリケーションと API の要件Office[を指定する」を参照してください](../develop/specify-office-hosts-and-api-requirements.md)。
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> コード内のランタイム チェックを使用して、ユーザーのホストとプラットフォームの組み合わせがこれらの要件セットをサポートするかどうかをテスト[](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)できます(「メソッドと要件セットのサポートのランタイム チェック」を参照)。 (マニフェストで要件セットを指定する方法は、この記事でも説明しますが、現在 RibbonApi 1.2 では機能しません)。または、カスタム コンテキスト タブがサポートされていない場合に、別の [UI エクスペリエンスを実装することもできます](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

## <a name="behavior-of-custom-contextual-tabs"></a>カスタム コンテキスト タブの動作

カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのコンテキスト タブのOfficeに従います。 配置のカスタム コンテキスト タブの基本的な原則を次に示します。

- カスタム コンテキスト タブが表示されている場合は、リボンの右側に表示されます。
- 1 つ以上の組み込みのコンテキスト タブと、アドインから 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側に表示されます。
- アドインに複数のコンテキスト タブが含み、複数のコンテキストが表示されている場合は、アドインで定義されている順序で表示されます。 (方向は Office 言語と同じ方向です。つまり、左から右の言語では左から右、右から左の言語では右から左)。定義[方法の詳細については、「タブに表示される](#define-the-groups-and-controls-that-appear-on-the-tab)グループとコントロールを定義する」を参照してください。
- 特定のコンテキストで表示されるコンテキスト タブが複数のアドインにある場合は、アドインが起動された順序で表示されます。
- カスタム *コンテキスト* タブは、カスタム コア タブとは異なり、アプリケーションのリボンOffice完全には追加されません。 アドインが実行されているOfficeドキュメントにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキスト タブを含む場合の主な手順

アドインにカスタム コンテキスト タブを含む場合の主な手順を次に示します。

1. 共有ランタイムを使用するアドインを構成します。
1. タブと、タブに表示されるグループとコントロールを定義します。
1. コンテキスト タブを [コンテキスト] タブにOffice。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

カスタム コンテキスト タブを追加するには、共有ランタイムを使用するアドインが必要です。 詳細については、「共有ランタイム [を使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェストで XML で定義されたカスタム コア タブとは異なり、カスタム コンテキスト タブは実行時に JSON BLOB を使用して定義されます。 コードは BLOB を JavaScript オブジェクトに解析し、オブジェクトを[Office.ribbon.requestCreateControls メソッドに渡](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)します。 カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。 これは、アドインのインストール時に Office アプリケーション リボンに追加されるカスタム コア タブとは異なります。また、別のドキュメントを開いた時点でも存在します。 また、 `requestCreateControls` メソッドはアドインのセッションで 1 回だけ実行できます。 再び呼び出された場合は、エラーがスローされます。

> [!NOTE]
> JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML の [CustomTab](../reference/manifest/customtab.md) 要素とその子孫要素の構造と大まかに平行です。

コンテキスト タブ JSON BLOB のステップ バイ ステップの例を作成します。 コンテキスト タブ JSON の完全なスキーマは [、dynamic-ribbon.schema.json にあります](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 このドキュメントで作業しているVisual Studio Code、このファイルを使用して、JSON を取得IntelliSense検証できます。 詳細については、「JSON スキーマと設定を使用Visual Studio Code JSON の編集[」を参照してください](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。

1. まず、という名前の 2 つの配列プロパティを持つ JSON 文字列を作成 `actions` します `tabs` 。 配列 `actions` は、コンテキスト タブのコントロールで実行できるすべての関数の仕様です。配列 `tabs` は、最大 *20* までの 1 つ以上のコンテキスト タブを定義します。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. コンテキスト タブのこの簡単な例では、ボタンが 1 つしか表示され、1 つのアクションだけが表示されます。 配列の唯一のメンバーとして次を追加 `actions` します。 このマークアップについて、次の点に注意してください。

    - プロパティ `id` と `type` プロパティは必須です。
    - 値には `type` 、"ExecuteFunction" または "ShowTaskpane" を指定できます。
    - プロパティ `functionName` は、 の値が . の場合にのみ `type` 使用されます `ExecuteFunction` 。 これは、FunctionFile で定義されている関数の名前です。 FunctionFile の詳細については、「アドイン コマンドの基本的 [な概念」を参照してください](add-in-commands.md)。
    - 後の手順では、このアクションをコンテキスト タブのボタンにマップします。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. 配列の唯一のメンバーとして次を追加 `tabs` します。 このマークアップについて、次の点に注意してください。

    - `id` プロパティは必須です。 アドイン内のすべてのコンテキスト タブで一意の簡潔でわかりやすい ID を使用します。
    - `label` プロパティは必須です。 コンテキスト タブのラベルとして機能するユーザーフレンドリーな文字列です。
    - `groups` プロパティは必須です。 タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバーと *20 以下のメンバーが必要です*。 (カスタム コンテキスト タブで使用できるコントロールの数にも制限があります。また、ユーザーが持つグループの数も制限されます。 詳細については、次の手順を参照してください)。

    > [!NOTE]
    > Tab オブジェクトには、アドインの起動時にタブをすぐに表示するかどうかを指定するオプション `visible` のプロパティを指定することもできます。 コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になります (ドキュメント内で何らかの種類のエンティティを選択するユーザーなど)、プロパティの既定値は存在しない場合です `visible` `false` 。 後のセクションでは、イベントに応答してプロパティを設定 `true` する方法を示します。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 単純な進行中の例では、コンテキスト タブには 1 つのグループのみがあります。 配列の唯一のメンバーとして次を追加 `groups` します。 このマークアップについて、次の点に注意してください。

    - すべてのプロパティが必要です。
    - プロパティ `id` は、マニフェスト内のすべてのグループ間で一意である必要があります。 最大 125 文字の簡潔でわかりやすい ID を使用します。
    - グループ `label` のラベルとして使用するユーザーフレンドリーな文字列です。
    - プロパティの値は、リボンのサイズとアプリケーション ウィンドウのサイズに応じて、グループがリボンに表示するアイコンをOffice `icon` 配列です。
    - プロパティ `controls` の値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。 少なくとも 1 つが必要です。

    > [!IMPORTANT]
    > *タブ全体のコントロールの総数は 20 以下です。* たとえば、各コントロールが 6 つのグループが 3 つ、コントロールが 2 つの 4 番目のグループを持つグループを 3 つ持つ場合がありますが、6 つのコントロールを持つグループを 4 つ持つ必要があります。  

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

1. すべてのグループには、32x32 px と 80x80 px の少なくとも 2 つのサイズのアイコンが必要です。 必要に応じて、サイズ 16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、および 64x64 px のアイコンを使用できます。 Office、リボンのサイズとアプリケーション ウィンドウのサイズに基づいて使用するOffice決定します。 アイコン配列に次のオブジェクトを追加します。 (ウィンドウとリボンのサイズが、グループのコントロールの少なくとも 1 つが表示されるのに十分な大きさの場合、グループ アイコンは表示されません。  たとえば、Word **ウィンドウを縮小** して展開する場合は、Word リボンの [スタイル] グループを確認します)。このマークアップについて、次の点に注意してください。

    - 両方のプロパティが必要です。
    - プロパティ `size` の単位はピクセルです。 アイコンは常に正方形なので、数値は高さと幅の両方です。
    - プロパティ `sourceLocation` は、アイコンの完全な URL を指定します。

    > [!IMPORTANT]
    > 開発から実稼働に移行する場合 (ドメインを localhost から contoso.com に変更するなど) ときに、アドインのマニフェスト内の URL を通常変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。

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

1. この単純な進行中の例では、グループにはボタンが 1 つのみです。 次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。 このマークアップについて、次の点に注意してください。

    - を除くすべての `enabled` プロパティが必要です。
    - `type` コントロールの種類を指定します。 値には、"Button"、"Menu"、または "MobileButton" を指定できます。
    - `id` 125 文字まで指定できます。
    - `actionId` は、配列で定義されているアクションの ID である必要 `actions` があります。 (このセクションの手順 1 を参照してください)。
    - `label` は、ボタンのラベルとして機能するユーザーフレンドリーな文字列です。
    - `superTip` は、豊富な形式のツール ヒントを表します。 プロパティと `title` プロパティ `description` の両方が必要です。
    - `icon` ボタンのアイコンを指定します。 グループ アイコンに関する前の説明もここでも適用されます。
    - `enabled` (省略可能) は、コンテキスト タブが表示されたら、ボタンを有効にするかどうかを指定します。 存在しない場合の既定値は `true` です。

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>requestCreateControls でコンテキスト タブを Officeに登録する

コンテキスト タブは[、Office.ribbon.requestCreateControls メソッドOffice呼](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)び出すことによって、ユーザーに登録されます。 これは通常、メソッドに割り当てられている関数またはメソッドで `Office.initialize` 行 `Office.onReady` われます。 これらのメソッドとアドインの初期化の詳細については、「Initialize [your Office アドイン」を参照してください](../develop/initialize-add-in.md)。 ただし、初期化後にいつでもメソッドを呼び出す場合があります。

> [!IMPORTANT]
> メソッド `requestCreateControls` は、アドインの特定のセッションで 1 回だけ呼び出される場合があります。 再び呼び出された場合は、エラーがスローされます。

次に例を示します。 JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があります。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>requestUpdate でタブが表示されるコンテキストを指定する

通常、ユーザーが開始したイベントがアドイン コンテキストを変更すると、カスタム コンテキスト タブが表示されます。 グラフ (ブックの既定のワークシート) がアクティブ化されている場合にのみ、タブを表示する必要があるシナリオExcel考えます。

まず、ハンドラーを割り当てる。 これは、一般的に、ハンドラー (後の手順で作成) をワークシート内のすべてのグラフのイベントに割り当てる次の例のようにメソッド `Office.onReady` `onActivated` `onDeactivated` で行われます。

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

次に、ハンドラーを定義します。 次に示すのは、単純な例ですが、関数のより堅牢なバージョンについては、この記事の後半の `showDataTab` [「HostRestartNeeded](#handle-the-hostrestartneeded-error) エラーの処理」を参照してください。 このコードについては、以下の点に注意してください。

- Office では、リボンの状態を更新するタイミングが制御されます。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestUpdate_input_)メソッドは、更新要求をキューに入れられます。 リボンが実際に更新される場合ではなく、要求をキューに入れ次第、メソッド `Promise` はオブジェクトを解決します。
- メソッドのパラメーターは `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) オブジェクトで、(1) *は JSON* で指定されたとおりにタブを ID で指定し、(2) はタブの表示を指定します。
- 同じコンテキストで表示するカスタム コンテキスト タブが複数ある場合は、配列に追加のタブ オブジェクトを追加 `tabs` します。

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

タブを非表示にするハンドラーは、プロパティをに設定する以外は、ほぼ `visible` 同じです `false` 。

JavaScript Officeには、オブジェクトを簡単に構築するためのいくつかのインターフェイス (型) も用意 `RibbonUpdateData` されています。 TypeScript の `showDataTab` 関数を次に示します。これらの型を使用します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効な状態を同時に切り替える

このメソッドは、カスタム コンテキスト タブまたはカスタム コア タブのカスタム ボタンの有効または無効の状態を切り替える `requestUpdate` 場合にも使用されます。この詳細については、「Enable [and Disable Add-in Commands」を参照してください](disable-add-in-commands.md)。 タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられます。 これは、単一の呼び出しで行います `requestUpdate` 。 次に、コンテキスト タブを表示すると同時に、コア タブのボタンが有効になっている例を示します。

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

次の例では、有効になっているボタンは、表示されるコンテキスト タブと全く同じです。

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

カスタム コンテキスト タブのボタンから作業ウィンドウを開く場合は、JSON でアクションを作成 `type` します `ShowTaskpane` 。 次に、アクションのプロパティ `actionId` を設定したボタン `id` を定義します。 これにより、マニフェストの Runtime 要素で **指定された既定** の作業ウィンドウが開きます。

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

既定の作業ウィンドウではない作業ウィンドウを開く場合は、アクションの定義で `sourceLocation` プロパティを指定します。 次の例では、別のボタンから 2 番目の作業ウィンドウを開きます。

> [!IMPORTANT]
>
> - アクションに `sourceLocation` a が指定されている場合、作業ウィンドウでは共有ランタイムは使用されません。 新しい JavaScript ランタイムで実行されます。
> - 共有ランタイムを使用できる作業ウィンドウは 1 つ以下なので、プロパティを省略できるアクションは 1 つ `ShowTaskpane` 以下 `sourceLocation` です。

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

渡される JSON BLOB は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法でローカライズされません (マニフェストからのローカライズの制御で `requestCreateControls` [説明します](../develop/localization.md#control-localization-from-the-manifest))。 代わりに、ローカライズは、ロケールごとに個別の JSON BLOB を使用して実行時に行う必要があります。 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用してください。 次に例を示します。

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

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>カスタム コンテキスト タブがサポートされていない場合に、別の UI エクスペリエンスを実装する

プラットフォーム、アプリケーション、Office、Officeの組み合わせはサポートされていません `requestCreateControls` 。 アドインは、これらの組み合わせの 1 つでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計されている必要があります。 次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。

#### <a name="use-noncontextual-tabs-or-controls"></a>コンテキスト以外のタブまたはコントロールを使用する

カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合、カスタム コンテキスト タブを実装するアドインでフォールバック エクスペリエンスを作成するように設計されたマニフェスト要素 [、OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)があります。

この要素を使用する最も簡単な方法は、アドインのカスタム コンテキスト タブのリボンカスタマイズを複製する 1 つ以上のカスタム コア タブ (つまり、非コンテキスト カスタム タブ) をマニフェストで定義する方法です。 ただし、重複するグループ、コントロール、およびメニュー Item 要素の最初の子要素として、カスタム コア `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` タブに追加します。 [](../reference/manifest/group.md) [](../reference/manifest/control.md)  その効果は次のとおりです。

- カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインが実行されている場合、カスタム コア グループとコントロールはリボンに表示されません。 代わりに、アドインがメソッドを呼び出す際に、カスタム コンテキスト タブが作成 `requestCreateControls` されます。
- アドインがサポートしないアプリケーションまたはプラットフォームで実行されている場合、要素はカスタム コア タブ `requestCreateControls` に表示されます。

次に例を示します。 カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに "MyButton" が表示されます。 ただし、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親グループとカスタム コア タブが表示されます。

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

その他の例については [、「OverriddenByRibbonApi」を参照してください](../reference/manifest/overriddenbyribbonapi.md)。

親グループまたはメニューにマークが付いている場合、そのグループは表示されません。カスタム コンテキスト タブがサポートされていない場合、すべての子マークアップは `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 無視されます。 したがって、これらの子要素の中に **OverriddenByRibbonApi** 要素がある場合や、その値が何かは関係ありません。 この意味は、メニュー項目またはコントロールをすべてのコンテキストで表示する必要がある場合、メニュー項目またはコントロールがマークされていない必要があるだけでなく、その親メニューとグループもこの方法でマークしなけらなければならないという意味です `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 

> [!IMPORTANT]
> グループまたはメニュー *のすべての子* 要素にマークを付けない `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 前の段落で指定した理由で親要素がマークされている場合、これは `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 意味をなします。 さらに、親の **OverriddenByRibbonApi** を削除 (またはに設定) すると、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親が表示されますが、サポートされている場合は空になります `false` 。 したがって、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親にマークを付け、 を指定します `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>指定したコンテキストで作業ウィンドウを表示または非表示にする API を使用する

**OverriddenByRibbonApi** の代わりに、アドインは、カスタム コンテキスト タブ上のコントロールの機能を複製する UI コントロールを使用して作業ウィンドウを定義できます。次に [、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__)メソッドと [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__)メソッドを使用して、コンテキスト タブがサポートされている場合に表示される作業ウィンドウを表示します。 これらのメソッドの使い方の詳細については、「アドインの作業ウィンドウを表示または非表示にするOffice[を参照してください](../develop/show-hide-add-in.md)。

### <a name="handle-the-hostrestartneeded-error"></a>HostRestartNeeded エラーの処理

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 コードでこのエラーを処理する必要があります。 次に、方法の例を示します。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

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
- Communityタブのサンプルのデモ

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]