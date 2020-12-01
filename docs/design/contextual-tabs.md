---
title: Office アドインでカスタムコンテキストタブを作成する
description: カスタムコンテキストタブを Office アドインに追加する方法について説明します。
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 49a773aca0651b88c972c24a4cde0aa1e300d5e7
ms.sourcegitcommit: 6619e07cdfa68f9fa985febd5f03caf7aee57d5e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/30/2020
ms.locfileid: "49505557"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Office アドインでカスタムコンテキストタブを作成する (プレビュー)

コンテキストタブは、office ドキュメントで指定されたイベントが発生したときにタブ行に表示される Office リボンの非表示タブコントロールです。 たとえば、テーブルが選択されているときに、Excel のリボンに表示される [ **テーブルデザイン** ] タブ。 Office アドインにカスタムコンテキストタブを含めることができ、表示または非表示のタイミングを指定するには、表示を変更するイベントハンドラーを作成します。 (ただし、カスタムコンテキストタブはフォーカスの変更には応答しません)。

> [!NOTE]
> この記事は、以下のドキュメントについて既に理解していることを前提としています。 最近、アドイン コマンド (カスタム メニュー項目とリボン ボタン) を使用してない場合は、ドキュメントをご確認ください。
>
> - [アドイン コマンドの基本概念](add-in-commands.md)

> [!IMPORTANT]
> ユーザー設定のコンテキストタブはプレビューで表示されます。 開発環境またはテスト環境で試してみることはできますが、運用アドインには追加しません。
>
> 現時点では、カスタムコンテキストタブは Excel でのみサポートされており、これらのプラットフォームでのみサポートされています。
>
> - Excel on Windows (Microsoft 365 のみ、永続的なライセンスではない): バージョン 2011 (ビルド 13426.20274)。 Microsoft 365 サブスクリプションは、以前は "月次 Channel (対象指定)" または "Insider 低速" と呼ばれていた [現在のチャネル (プレビュー)](https://insider.office.com/join/windows) にある必要があります。

> [!NOTE]
> カスタムコンテキストタブは、次の要件セットをサポートするプラットフォームでのみ機能します。 要件セットの詳細とその使用方法については、「 [Office アプリケーションと API 要件を指定](../develop/specify-office-hosts-and-api-requirements.md)する」を参照してください。
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>カスタムコンテキストタブの動作

カスタムコンテキストタブのユーザー操作は、組み込みの Office コンテキストタブのパターンに従います。 配置カスタムコンテキストタブの基本的な原則を次に示します。

- ユーザー設定のコンテキストタブが表示されている場合は、リボンの右端に表示されます。
- 1つ以上の組み込みコンテキストタブと、アドインのカスタムコンテキストタブが同時に表示されている場合は、カスタムコンテキストタブは常に、組み込みのコンテキストタブすべての右側にあります。
- アドインに複数のコンテキストタブがあり、複数のコンテキストタブが表示されている場合は、アドインで定義された順序で表示されます。 (方向は、Office の言語と同じ方向です。つまり、左から右に記述する言語では左から右ですが、右から左の言語では右から左)。定義方法の詳細については、 [タブに表示されるグループとコントロールを定義](#define-the-groups-and-controls-that-appear-on-the-tab) するを参照してください。
- 複数のアドインにコンテキストタブがあり、特定のコンテキストに表示されている場合は、アドインが起動された順序で表示されます。
- カスタムの *コンテキスト* タブは、カスタムコアタブとは異なり、Office アプリケーションのリボンに永続的に追加されません。 これらは、アドインが実行されている Office ドキュメントにのみ存在します。

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>アドインにコンテキストタブを含めるための主な手順

アドインにカスタムコンテキストタブを含めるための主な手順を次に示します。

1. 共有ランタイムを使用するようにアドインを構成します。
1. タブと、その上に表示されるグループとコントロールを定義します。
1. コンテキストタブを Office に登録します。
1. タブが表示される状況を指定します。

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

カスタムコンテキストタブを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>タブに表示されるグループとコントロールを定義する

マニフェストの XML で定義されているカスタムコアタブとは異なり、カスタムコンテキストタブは、JSON blob を使用して実行時に定義されます。 コードによって blob が JavaScript オブジェクトに解析され、そのオブジェクトが [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) メソッドに渡されます。 カスタムコンテキストタブは、アドインが現在実行されているドキュメントにのみ存在します。 これは、アドインがインストールされていて、別のドキュメントが開かれたときにそのまま表示される場合に、Office アプリケーションリボンに追加されるカスタムコアタブとは異なります。 また、この `requestCreateControls` メソッドは、アドインのセッションで一度だけ実行できます。 再び呼び出された場合は、エラーがスローされます。

> [!NOTE]
> JSON blob のプロパティとサブプロパティ (およびキー名) の構造は、マニフェスト XML 内の [Customtab](../reference/manifest/customtab.md) 要素とその子孫要素の構造とほぼ並行しています。

ここでは、コンテキストタブ JSON blob の詳細な手順を示します。 (コンテキストタブ JSON の完全なスキーマは [dynamic-ribbon.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)あります。 このリンクは、コンテキストタブの初期プレビュー期間では動作していない場合があります。 リンクが機能していない場合は、 [ドラフト dynamic-ribbon.schema.js](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)のスキーマの最新の下書きを見つけることができます。)Visual Studio Code で作業している場合は、このファイルを使用して IntelliSense を取得し、JSON を検証することができます。 詳細については、「 [Visual Studio Code-json スキーマと設定を使用](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)して Json を編集する」を参照してください。


1. 最初に、という名前の2つの配列プロパティを使用して JSON 文字列を作成し `actions` `tabs` ます。 配列は、 `actions` コンテキストタブのコントロールによって実行できるすべての関数の仕様です。配列は、 `tabs` 1 つ以上のコンテキストタブを定義します。 *最大値は 10* です。

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. このような状況に応じたタブの例では、1つのボタンだけが含まれるため、1つのアクションのみができます。 次のものを配列の唯一のメンバーとして追加し `actions` ます。 このマークアップについて、次の点に注意してください。

    - `id`プロパティおよび `type` プロパティは必須です。
    - の値は、 `type` "ExecuteFunction" または "ShowTaskpane" のいずれかになります。
    - `functionName`プロパティは、の値がである場合にのみ使用され `type` `ExecuteFunction` ます。 これは、FunctionFile で定義されている関数の名前です。 FunctionFile の詳細については、「 [アドインコマンドの基本的な概念](add-in-commands.md)」を参照してください。
    - この後の手順では、このアクションをコンテキストタブのボタンにマップします。

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. 次のものを配列の唯一のメンバーとして追加し `tabs` ます。 このマークアップについて、次の点に注意してください。

    - `id` プロパティは必須です。 簡単でわかりやすい ID を使用して、アドイン内のすべてのコンテキストタブにおいて一意にします。
    - `label` プロパティは必須です。 これは、コンテキストタブのラベルとして機能する、ユーザーフレンドリな文字列です。
    - `groups` プロパティは必須です。 タブに表示されるコントロールのグループを定義します。少なくとも1つのメンバーがあり *、20を超える* ことはできません。 (カスタムコンテキストタブで使用できるコントロールの数にも制限があります。また、グループの数も制限されます。 詳細については、次の手順を参照してください)。

    > [!NOTE]
    > Tab オブジェクトには、 `visible` アドインの起動時にすぐにタブを表示するかどうかを指定するオプションのプロパティもあります。 通常、コンテキストタブは、ユーザーイベントによって表示がトリガーされる (ドキュメント内の一部の型のエンティティを選択するユーザーのような場合) ため、 `visible` プロパティは `false` 存在しない場合は既定値になります。 後のセクションでは、イベントへの応答としてプロパティをに設定する方法を示し `true` ます。

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. この単純な例では、コンテキストタブには1つのグループしかありません。 次のものを配列の唯一のメンバーとして追加し `groups` ます。 このマークアップについて、次の点に注意してください。

    - すべてのプロパティが必要です。
    - この `id` プロパティは、タブ内のすべてのグループ間で一意である必要があります。簡潔でわかりやすい ID を使用してください。
    - `label`は、グループのラベルとして機能する、ユーザーフレンドリな文字列です。
    - `icon`プロパティの値は、リボンと Office アプリケーションウィンドウのサイズに応じて、グループがリボン上に持つアイコンを指定するオブジェクトの配列です。
    - `controls`プロパティの値は、グループ内のボタンやその他のコントロールを指定するオブジェクトの配列です。 *グループに* は、少なくとも1つの値が必要です。

    > [!IMPORTANT]
    > *タブ全体のコントロールの合計数は20以下でなければなりません。* たとえば、6つのグループと2つのコントロールを持つ4つのグループを持つことができますが、それぞれ6つのコントロールを持つ4つのグループを持つことはできません。  

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

1. すべてのグループには、少なくとも2つのサイズのアイコン、32x32 px、80x80 px が必要です。 必要に応じて、サイズが16x16、20x20、24x24、40x40、48x48、64x64 のアイコンを設定することもできます。 Office は、リボンおよび Office アプリケーションウィンドウのサイズに基づいて、どのアイコンを使用するかを決定します。 アイコン配列に次のオブジェクトを追加します。 (ウィンドウ上の少なくとも1つの *コントロール* が表示されるようにウィンドウとリボンのサイズが大きい場合、[グループ] アイコンはまったく表示されません。 例については、word のウィンドウを縮小して展開するときに、Word のリボンの [ **スタイル** ] グループを見てください。このマークアップについて、次の点に注意してください。

    - 両方のプロパティが必要です。
    - `size`プロパティの測定単位はピクセルです。 アイコンは常に正方形なので、数字は高さと幅の両方です。
    - この `sourceLocation` プロパティは、アイコンへの完全な URL を指定します。

    > [!IMPORTANT]
    > 開発環境から運用環境に移行するときに、通常、アドインのマニフェストの Url を変更する必要があるのと同様に、コンテキストタブ JSON の Url も変更する必要があります。

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

1. この簡単な例では、グループにボタンが1つしかありません。 次のオブジェクトを配列の唯一のメンバーとして追加し `controls` ます。 このマークアップについて、次の点に注意してください。

    - 以外のすべてのプロパティ `enabled` は必須です。
    - `type` コントロールの種類を指定します。 値には、"Button"、"Menu"、または "MobileButton" を指定できます。
    - `id` 最大125文字を使用できます。 
    - `actionId` は、配列で定義されているアクションの ID である必要があり `actions` ます。 (このセクションの手順1を参照してください)。
    - `label` は、ボタンのラベルとして機能するわかりやすい文字列です。
    - `superTip` ツールヒントの豊富な形式を表します。 `title`プロパティとプロパティの両方 `description` が必要です。
    - `icon` ボタンのアイコンを指定します。 グループアイコンに関する上記の解説も適用されます。
    - `enabled` (省略可能) [操作] タブが表示されたときにボタンを有効にするかどうかを指定します。 存在しない場合の既定値は、 `true` です。 

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
 
JSON blob の完全な例を次に示します。

```json
'{
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
      "label": "Data",
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
}'
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>RequestCreateControls を使用して Office にコンテキストタブを登録する

コンテキストタブは、 [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) メソッドを呼び出すことによって、office に登録されています。 これは、通常、またはメソッドによって割り当てられた関数によって実行され `Office.initialize` `Office.onReady` ます。 これらのメソッドとアドインの初期化の詳細については、「 [Office アドインを初期化する](../develop/initialize-add-in.md)」を参照してください。 ただし、初期化後はいつでもメソッドを呼び出すことができます。

> [!IMPORTANT]
> この `requestCreateControls` メソッドは、アドインの特定のセッションで1回だけ呼び出すことができます。 再び呼び出された場合、エラーがスローされます。

次に例を示します。 JSON 文字列は、 `JSON.parse` javascript 関数に渡す前に、メソッドを使用して javascript オブジェクトに変換する必要があります。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>RequestUpdate でタブが表示されるときのコンテキストを指定する

通常、ユーザーが開始したイベントによってアドインコンテキストが変更されたときに、ユーザー設定のコンテキストタブが表示されるようにする必要があります。 Excel ブックの既定のワークシートにあるグラフがアクティブ化されたときにのみタブが表示されるシナリオを考えてみます。

最初に、ハンドラーを割り当てます。 これは通常、次の例のようにメソッドで行われ `Office.onReady` ます。この例では、ハンドラー (後の手順で作成したもの) を、 `onActivated` `onDeactivated` ワークシート内のすべてのグラフのイベントに割り当てます。

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

次に、ハンドラーを定義します。 の簡単な例を次に示し `showDataTab` ますが、より堅牢なバージョンの関数については、この記事で後述する「 [エラー処理](#error-handling) 」を参照してください。 このコードについては、以下の点に注意してください。

- Office では、リボンの状態を更新するタイミングが制御されます。 [Office の更新](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)要求をキューに入れます。 このメソッドは、 `Promise` リボンが実際に更新されるときではなく、要求がキューに入った直後にオブジェクトを解決します。
- メソッドのパラメーター `requestUpdate` は、 *JSON で指定され* た ID を使用してタブを指定する [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata)オブジェクトであり、(2) はタブの表示を指定します。
- 同じコンテキストに表示する必要があるカスタムコンテキストタブが複数ある場合は、単に配列に追加の tab オブジェクトを追加し `tabs` ます。

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

タブを非表示にするハンドラーはほぼ同じですが、プロパティがに戻される点が異なり `visible` `false` ます。

Office JavaScript ライブラリには、オブジェクトを簡単に作成できるように、いくつかのインターフェイス (型) も用意されて `RibbonUpdateData` います。 次に示すのは `showDataTab` TypeScript の関数で、これらの型を使用します。

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>ボタンの表示/非表示の状態を同時に切り替えます。

この `requestUpdate` メソッドは、カスタムコンテキストタブまたはカスタムコアタブのいずれかで、カスタムボタンの有効または無効の状態を切り替えるためにも使用されます。詳細については、「 [アドインコマンドを有効または無効](disable-add-in-commands.md)にする」を参照してください。 タブの表示とボタンの有効な状態の両方を同時に変更するシナリオがある場合があります。 これは、の1回の呼び出しで行うことができ `requestUpdate` ます。 次の例は、コンテキストタブが表示されているときに、コアタブのボタンが有効になっています。

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

次の例では、有効になっているボタンは、表示されているのと同じコンテキストタブにあります。

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="error-handling"></a>エラー処理

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
