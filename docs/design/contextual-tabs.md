---
title: Office アドインでカスタム コンテキスト タブを作成する
description: Office アドインにカスタム コンテキスト タブを追加する方法について説明します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d03ac2c01c03353f3e2d1b54ba20616d7b42d93f
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555207"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Office アドインでカスタム コンテキスト タブを作成する

コンテキスト タブは、Officeのリボンの非表示のタブ コントロールで、指定したイベントがOffice ドキュメントで発生したときにタブ行に表示されます。 たとえば、テーブルが選択されたときにリボンExcel表示される [テーブル **デザイン**] タブなどです。 Office アドインにカスタム コンテキスト タブを含め、表示を変更するイベント ハンドラーを作成して、表示または非表示を切り替えるタイミングを指定できます。 (ただし、カスタム コンテキスト タブはフォーカスの変更に応答しません)。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

> [!IMPORTANT]
> カスタム コンテキスト タブは現在、Excelでのみサポートされており、これらのプラットフォームとビルドでのみサポートされています。
>
> - Windows (Microsoft 365 サブスクリプションのみ) でExcel): バージョン 2102 (ビルド 13801.20294) 以降。
> - Excel on the web

> [!NOTE]
> カスタム コンテキスト タブは、次の要件セットをサポートするプラットフォームでのみ機能します。 要件セットの詳細と、それらの要件セットの使用方法については[、「Officeアプリケーションと API 要件の指定](../develop/specify-office-hosts-and-api-requirements.md)」を参照してください。
>
> - [リボンアピ1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> コードでランタイム チェックを使用して、ユーザーのホストとプラットフォームの組み合わせがこれらの要件[Office](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)セットをサポートしているかどうかをテストできます 。 (マニフェストで要件セットを指定する手法は、その記事でも説明されていますが、現在のところ、RibbonApi 1.2 では機能しません)。または、 [カスタム コンテキスト タブがサポートされていない場合に、代替 UI エクスペリエンスを実装](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)することもできます。

## <a name="behavior-of-custom-contextual-tabs"></a>カスタム コンテキスト タブの動作

カスタム コンテキスト タブのユーザー エクスペリエンスは、組み込みのOfficeコンテキスト タブのパターンに従います。 配置カスタム コンテキスト タブの基本原則を次に示します。

- カスタム コンテキスト タブが表示されている場合、リボンの右端に表示されます。
- 1 つ以上の組み込みコンテキスト タブと、アドインの 1 つ以上のカスタム コンテキスト タブが同時に表示される場合、カスタム コンテキスト タブは常にすべての組み込みコンテキスト タブの右側にあります。
- アドインに複数のコンテキスト タブがあり、複数のコンテキストが表示されている場合、アドインで定義されている順序で表示されます。 (方向はOffice言語と同じ方向、つまり、左から右に右に、右から左の言語では右から左の言語です。定義方法の詳細については、「[タブに表示されるグループとコントロール](#define-the-groups-and-controls-that-appear-on-the-tab)の定義」を参照してください。
- 複数のアドインに特定のコンテキストで表示されるコンテキスト タブがある場合、アドインが起動された順序で表示されます。
- カスタム *のコンテキスト* タブは、カスタム コア タブとは異なり、Office アプリケーションのリボンに永続的に追加されません。 これらのファイルは、アドインが実行されているドキュメントOfficeにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキスト タブを含める主な手順

アドインにカスタム コンテキスト タブを含める主な手順を次に示します。

1. 共有ランタイムを使用するようにアドインを構成します。
1. タブ、およびタブに表示されるグループとコントロールを定義します。
1. コンテキスト タブをOfficeに登録します。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

カスタム コンテキスト タブを追加するには、共有ランタイムを使用するアドインが必要です。 詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェストで XML で定義されるカスタム コア タブとは異なり、カスタム コンテキスト タブは JSON BLOB を使用して実行時に定義されます。 コードは、Blob を JavaScript オブジェクトに解析し、そのオブジェクトを[Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)メソッドに渡します。 カスタム コンテキスト タブは、アドインが現在実行されているドキュメントにのみ表示されます。 これは、アドインがインストールされるときにOffice アプリケーション リボンに追加され、別のドキュメントが開かれたときに表示されたままになるカスタム コア タブとは異なります。 また、 `requestCreateControls` このメソッドは、アドインのセッションで 1 回だけ実行できます。 再度呼び出されると、エラーがスローされます。

> [!NOTE]
> JSON BLOB のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML の [CustomTab](../reference/manifest/customtab.md) 要素とその子孫要素の構造とほぼ平行です。

コンテキスト タブ JSON BLOB の例を段階的に作成します。 コンテキスト タブ JSON の完全なスキーマは、 [dynamic-ribbon.schema.jsにあります](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。 Visual Studio Codeで作業している場合は、このファイルを使用してIntelliSenseを取得し、JSON を検証できます。 詳細については[、「json の編集 」を参照してくださいVisual Studio Code - JSON スキーマと設定を](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)使用します。


1. まず、 という名前の 2 つの配列プロパティを持つ JSON 文字列 `actions` を作成 `tabs` します。 `actions`配列は、コンテキスト タブのコントロールで実行できるすべての関数の仕様です。`tabs`配列は、*最大 20* 個までの 1 つ以上のコンテキスト タブを定義します。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. この単純なコンテキスト タブの例では、ボタンは 1 つだけで、単一のアクションのみが表示されます。 次の要素を配列の唯一のメンバーとして追加 `actions` します。 このマークアップについて、次の点に注意してください。

    - `id`および `type` プロパティは必須です。
    - の値 `type` は、"関数の実行" または "タスク ウィンドウの表示" のいずれかです。
    - `functionName`プロパティは、 の値が の場合にのみ使用 `type` されます `ExecuteFunction` 。 これは、関数ファイルで定義されている関数の名前です。 FunctionFile の詳細については、「アドイン [コマンドの基本概念](add-in-commands.md)」を参照してください。
    - 後の手順で、このアクションをコンテキスト タブのボタンにマップします。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. 次の要素を配列の唯一のメンバーとして追加 `tabs` します。 このマークアップについて、次の点に注意してください。

    - `id` プロパティは必須です。 アドイン内のすべてのコンテキスト タブに固有の簡単な説明 ID を使用します。
    - `label` プロパティは必須です。 コンテキスト タブのラベルとして使用するわかりやすい文字列です。
    - `groups` プロパティは必須です。 タブに表示されるコントロールのグループを定義します。少なくとも 1 つのメンバー *と 20 以下の メンバーが* 必要です。 (カスタム コンテキスト タブに設定できるコントロールの数にも制限があり、グループの数も制限されます。 詳細については、次の手順を参照してください。

    > [!NOTE]
    > Tab オブジェクトには、 `visible` アドインの起動時にタブをすぐに表示するかどうかを指定するオプションのプロパティを持つ場合もあります。 コンテキスト タブは通常、ユーザー イベントが表示をトリガーするまで非表示になるため (ユーザーがドキュメント内の何らかの種類のエンティティを選択した場合など)、 `visible` プロパティは既定で `false` 表示されない場合に設定されます。 後のセクションでは、イベントに応答してプロパティを設定 `true` する方法を示します。

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. 簡単な進行中の例では、コンテキスト タブには 1 つのグループしかありません。 次の要素を配列の唯一のメンバーとして追加 `groups` します。 このマークアップについて、次の点に注意してください。

    - すべてのプロパティが必要です。
    - `id`プロパティは、タブ内のすべてのグループ間で一意である必要があります。
    - `label`は、グループのラベルとして使用するわかりやすい文字列です。
    - `icon`プロパティの値は、リボンのサイズとアプリケーション ウィンドウに応じてリボンにグループが表示されるアイコンを指定するオブジェクトの配列Office。
    - `controls`プロパティの値は、グループ内のボタンとメニューを指定するオブジェクトの配列です。 少なくとも 1 つ存在する必要があります。

    > [!IMPORTANT]
    > *タブ全体のコントロールの合計数は 20 個以下にできます。* たとえば、それぞれ 6 つのコントロールを持つ 3 つのグループと、4 番目のグループに 2 つのコントロールを含め、それぞれ 6 つのコントロールを持つ 4 つのグループを持つことはできません。  

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

1. 各グループには、32x32 pxと80x80 pxの2つ以上のサイズのアイコンが必要です。 オプションで、サイズ 16x16 px、20x20 px、24x24 px、40x40 px、48x48 px、および 64x64 ピクセルのアイコンを使用することもできます。 Office、リボンのサイズとアプリケーション ウィンドウに基づいて、使用するアイコンOffice決定します。 アイコン配列に次のオブジェクトを追加します。 (ウィンドウとリボンのサイズが *グループのコントロール* の少なくとも 1 つが表示されるのに十分な大きさの場合は、グループ アイコンはまったく表示されません。 たとえば、Word ウィンドウを縮小および展開するときに、Word リボンの **[スタイル]** グループを確認します。このマークアップについて、次の点に注意してください。

    - 両方のプロパティが必要です。
    - `size`プロパティの単位はピクセルです。 アイコンは常に正方形であるため、数値は高さと幅の両方になります。
    - `sourceLocation`プロパティは、アイコンへの完全な URL を指定します。

    > [!IMPORTANT]
    > 開発から運用環境に移行する場合 (localhost から contoso.com にドメインを変更するなど)、アドインのマニフェストの URL を通常変更する必要があるのと同様に、コンテキスト タブ JSON の URL も変更する必要があります。

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

1. 簡単な進行中の例では、グループにはボタンが 1 つしかありません。 次のオブジェクトを配列の唯一のメンバーとして追加 `controls` します。 このマークアップについて、次の点に注意してください。

    - を除くすべてのプロパティ `enabled` が必要です。
    - `type` コントロールの種類を指定します。 値は"ボタン"、"メニュー"、または"モバイルボタン"にすることができます。
    - `id` 最大 125 文字まで指定できます。 
    - `actionId` は、配列で定義されたアクションの ID でなければなりません `actions` 。 (このセクションのステップ 1 を参照してください。
    - `label` は、ボタンのラベルとして使用するわかりやすい文字列です。
    - `superTip` は、ツール ヒントの豊富な形式を表します。 `title`プロパティと プロパティ `description` の両方が必要です。
    - `icon` ボタンのアイコンを指定します。 グループアイコンに関する以前の解説もここに当てはまります。
    - `enabled` (オプション)は、コンテキストタブが表示されたときにボタンを有効にするかどうかを指定します。 存在しない場合のデフォルトは `true` です。 

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>コンテキスト タブをOfficeに登録します。

コンテキスト タブは[、Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)メソッドを呼び出すことによってOfficeに登録されます。 これは通常、メソッドに割り当てられている関数 `Office.initialize` またはメソッドを使用して実行されます `Office.onReady` 。 これらのメソッドとアドインの初期化の詳細については[、「Office アドインの初期化」を](../develop/initialize-add-in.md)参照してください。 ただし、初期化後はいつでもメソッドを呼び出すことができます。

> [!IMPORTANT]
> `requestCreateControls`このメソッドは、アドインの特定のセッションで 1 回だけ呼び出すことができます。 再び呼び出されると、エラーがスローされます。

次に例を示します。 JSON 文字列を JavaScript 関数に渡す前に、メソッドを使用して `JSON.parse` JavaScript オブジェクトに変換する必要があることに注意してください。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>タブが requestUpdate で表示されるコンテキストを指定します。

通常、ユーザーが開始したイベントがアドインコンテキストを変更したときに、カスタム コンテキスト タブが表示されます。 グラフ (Excel ブックの既定のワークシート) がアクティブになったときに、タブを表示する必要があるシナリオを考えてみます。

まず、ハンドラを割り当てます。 これは、 `Office.onReady` 通常、このメソッドで、後の手順で作成したハンドラーを、 ワークシート `onActivated` 内のすべてのグラフの イベントと に割り当てる方法 `onDeactivated` と同様に行われます。

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

次に、ハンドラーを定義します。 次に示す簡単な例を `showDataTab` 示しますが、関数のより堅牢なバージョンについては、この記事の後の [「HostRestartNeeded エラーの処理](#handle-the-hostrestartneeded-error) 」を参照してください。 このコードについては、以下の点に注意してください。

- Office では、リボンの状態を更新するタイミングが制御されます。 [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)メソッドは、更新要求をキューに入れます。 このメソッドは、 `Promise` リボンが実際に更新されたときではなく、要求をキューに入れるとすぐにオブジェクトを解決します。
- メソッドのパラメーター `requestUpdate` は、(1) *JSON で指定されたとおり* にタブを ID で指定し、(2) タブの可視性を指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata)オブジェクトです。
- 同じコンテキストで表示する必要があるカスタム コンテキスト タブが複数ある場合は、単に配列にタブ オブジェクトを追加するだけです `tabs` 。

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

タブを非表示にするハンドラーはほぼ同じですが、プロパティを `visible` に戻す点が異なっています `false` 。

JavaScript ライブラリOfficeには、オブジェクトの構築を容易にするためにいくつかのインターフェイス (型) も用意 `RibbonUpdateData` されています。 TypeScript の `showDataTab` 関数を次に示します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>タブの表示とボタンの有効ステータスを同時に切り替える

この `requestUpdate` メソッドは、カスタム コンテキスト タブまたはカスタム コア タブでカスタム ボタンの有効または無効の状態を切り替える場合にも使用されます。詳細については、「 アドイン [コマンドの有効化と無効化](disable-add-in-commands.md)」を参照してください。 タブの表示とボタンの有効な状態の両方を同時に変更するシナリオが考えられる場合があります。 これは、 の 1 回の呼び出しで行うことができます `requestUpdate` 。 次の例では、コンテキスト タブが表示されるようにすると同時に、コア タブのボタンが有効になります。

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

次の例では、有効になっているボタンは、表示されているコンテキスト タブとまったく同じです。

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

渡される JSON BLOB `requestCreateControls` は、カスタム コア タブのマニフェスト マークアップがローカライズされるのと同じ方法でローカライズされません ( [これは、「マニフェストからのコントロールのローカライズ](../develop/localization.md#control-localization-from-the-manifest)」で説明しています)。 代わりに、ロケールごとに異なる JSON BLOB を使用して、実行時にローカリゼーションを実行する必要があります。 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)プロパティをテストするステートメントを使用することをお勧めします。 例を次に示します。

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

次に、次の例のように、コードは関数を呼び出して、 に渡されるローカライズされた BLOB `requestCreateControls` を取得します。

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>カスタム コンテキスト タブのベスト プラクティス

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>カスタム コンテキスト タブがサポートされていない場合に、代替 UI エクスペリエンスを実装する

プラットフォーム、Office アプリケーション、およびOffice ビルドの一部の組み合わせでは、 がサポートしていません `requestCreateControls` 。 アドインは、これらの組み合わせのいずれかでアドインを実行しているユーザーに代替エクスペリエンスを提供するように設計する必要があります。 次のセクションでは、フォールバック エクスペリエンスを提供する 2 つの方法について説明します。

#### <a name="use-noncontextual-tabs-or-controls"></a>非コンテキスト タブまたはコントロールを使用する

カスタム コンテキスト タブをサポートしないアプリケーションまたはプラットフォームでアドインが実行されている場合に、カスタム コンテキスト タブを実装するアドインでフォールバック エクスペリエンスを作成するように設計されたマニフェスト要素 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)があります。 

この要素を使用する最も簡単な方法は、アドインのカスタム コンテキスト タブのリボンのカスタマイズを複製する 1 つ以上のカスタム コア タブ ( *つまり、非コンテキスト* カスタム タブ) をマニフェストで定義することです。 しかし、 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` [あなたは CustomTab](../reference/manifest/customtab.md)の最初の子要素として追加します。 その結果、次のようになります。

- カスタム コンテキスト タブをサポートするアプリケーションとプラットフォームでアドインを実行する場合、カスタム コア タブはリボンに表示されません。 代わりに、アドインがメソッドを呼び出したときにカスタム コンテキスト タブが作成されます `requestCreateControls` 。
- アドインが をサポートしていないアプリケーションまたはプラットフォームで実行されている場合 *、* カスタム `requestCreateControls` コア タブがリボンに表示されます。

この単純な戦略の例を次に示します。

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

この単純な戦略では、カスタム コンテキスト タブを子グループとコントロールと共に反映するカスタム コア タブを使用しますが、より複雑な戦略を使用できます。 `<OverriddenByRibbonApi>`要素は[、(](../reference/manifest/group.md)最初の) 子要素として 、グループ要素および[コントロール](../reference/manifest/control.md)要素 ([ボタンの種類](../reference/manifest/control.md#button-control)と[メニューの種類](../reference/manifest/control.md#menu-dropdown-button-controls)の両方) およびメニュー要素に追加することもできます `<Item>` 。 この事実により、コンテキスト タブに表示されるグループやコントロールを、さまざまなカスタム コア タブのグループ、ボタン、メニューに分散させることができます。 次に例を示します。 カスタム コンテキスト タブがサポートされていない場合にのみ、カスタム コア タブに "MyButton" が表示されることに注意してください。 ただし、カスタム コンテキスト タブがサポートされているかどうかに関係なく、親グループとカスタム コア タブが表示されます。

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

その他の例については、「 [オーバーライドされた ByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)」を参照してください。

親タブ、グループ、またはメニューに マークが付いている場合、 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` そのタブは表示されず、カスタム コンテキスト タブがサポートされていない場合は、子マークアップはすべて無視されます。 したがって、これらの子要素のいずれかが要素を持 `<OverriddenByRibbonApi>` っているかどうか、またはその値は関係ありません。 このことは、メニュー項目、コントロール、またはグループがすべてのコンテキストで表示される必要がある場合、そのメニュー項目、コントロール、またはグループをでマークしないだけでなく `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *、その親メニュー、グループ、およびタブもこのようにマークしてはならない* ということです。

> [!IMPORTANT]
> タブ、グループ、または *メニューのすべての子* 要素を にマークを付けないでください `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 前の段落で指定した理由で親要素にマークが付いている場合、これは無意味です `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。 さらに、親の を除外する `<OverriddenByRibbonApi>` (または に設定 `false` する) 場合、カスタム コンテキスト タブがサポートされているかどうかに関係なく親が表示されますが、サポートされている場合は空になります。 したがって、カスタム コンテキスト タブがサポートされているときにすべての子要素が表示されない場合は、親と親のみを `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` でマークします。

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>指定したコンテキストで作業ウィンドウの表示と非表示を切り替える API を使用する

アドインの代わりに `<OverriddenByRibbonApi>` 、カスタム コンテキスト タブのコントロールの機能を複製する UI コントロールを含む作業ウィンドウを定義できます。次に[、Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__)メソッドと[Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__)メソッドを使用して、サポートされている場合にコンテキスト タブが表示された場合、およびコンテキスト タブが表示された場合にのみ、作業ウィンドウを表示します。 これらの方法の詳細については、「 [Office アドインの作業ウィンドウの表示と非表示を切り替える](../develop/show-hide-add-in.md)」を参照してください。

### <a name="handle-the-hostrestartneeded-error"></a>ホスト再起動が必要なエラーを処理します。

一部のシナリオでは、Office はリボンを更新できず、エラーを返します。 たとえば、アドインがアップグレードされ、アップグレードされたアドインに異なるカスタム アドイン コマンドのセットがある場合は、Office アプリケーションを閉じてから、もう一度開く必要があります。 それまでの間、`requestUpdate` メソッドは `HostRestartNeeded` エラーを返します。 コードでこのエラーを処理する必要があります。 以下は、その方法の例です。 この場合、`reportError` メソッドがユーザーにエラーを表示します。

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
