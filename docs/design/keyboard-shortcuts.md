---
title: Office アドインのカスタム キーボード ショートカット
description: カスタム キーボード ショートカット (キーの組み合わせとも呼ばれる) をアドインに追加Office説明します。
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: dc99674b92ebb415b1d49fb28821d8c2e34c8077
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789150"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>カスタム キーボード ショートカットを Office アドインに追加する (プレビュー)

キーボード ショートカット (キーの組み合わせとも呼ばれる) を使用すると、アドインのユーザーの作業効率が向上し、マウスに代わる方法が提供され、障がいのあるユーザーのアドインのアクセシビリティが向上します。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> キーボード ショートカットが既に有効になっているアドインの作業バージョンから開始するには、 [サンプルの Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)キーボード ショートカットを複製して実行します。 独自のアドインにキーボード ショートカットを追加する準備ができたら、この記事に進む必要があります。

アドインにキーボード ショートカットを追加するには、次の 3 つの手順があります。

1. [アドインのマニフェストを構成します](#configure-the-manifest)。
1. [ショートカットの JSON ファイルを作成または編集して](#create-or-edit-the-shortcuts-json-file) 、アクションとそのキーボード ショートカットを定義します。
1. [](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate API の 1 つ以上](/javascript/api/office/office.actions#associate)のランタイム呼び出しを追加して、各アクションに関数をマップします。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストには 2 つの小さな変更があります。 1 つは、アドインで共有ランタイムを使用し、もう 1 つは、キーボード ショートカットを定義した JSON 形式のファイルを指し示す方法です。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

カスタム キーボード ショートカットを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、共有 [ランタイムを使用するアドインを構成します](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>マッピング ファイルをマニフェストにリンクする

マニフェスト *内* の (内部ではなく) 要素の直下に `<VersionOverrides>` [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。 この属性は、後の手順で作成するプロジェクトの JSON ファイルの `Url` 完全な URL に設定します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>ショートカットの JSON ファイルを作成または編集する

プロジェクトに JSON ファイルを作成します。 ファイルのパスが ExtendedOverrides 要素の属性に指定した場所と一致 `Url` [する必要](../reference/manifest/extendedoverrides.md) があります。 このファイルには、キーボード ショートカットと、キーボード ショートカットが呼び出す操作が記述されます。

1. JSON ファイル内には 2 つの配列があります。 アクション配列には、呼び出されるアクションを定義するオブジェクトが含まれます。また、ショートカット配列には、キーの組み合わせをアクションにマップするオブジェクトが含まれます。 次に例を示します：

    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    JSON オブジェクトの詳細については、「アクション[](#constructing-the-action-objects)オブジェクトの構築」および「ショートカット オブジェクトの作成」[を参照してください](#constructing-the-shortcut-objects)。 ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

    > [!NOTE]
    > この記事では、"CTRL" の代えに "CONTROL" を使用できます。

    後の手順では、アクション自体が、作成する関数にマップされます。 この例では、後で SHOWTASKPANE をメソッドを呼び出す関数にマップし、HIDETASKPANE をメソッドを呼び出す `Office.addin.showAsTaskpane` 関数にマップ `Office.addin.hide` します。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>アクションの関数へのマッピングを作成する

1. プロジェクトで、HTML ページによって読み込まれた JavaScript ファイルを要素で開 `<FunctionFile>` きます。
1. JavaScript ファイルで [、Office.actions.associate](/javascript/api/office/office.actions#associate) API を使用して、JSON ファイルで指定した各アクションを JavaScript 関数にマップします。 次の JavaScript をファイルに追加します。 コードについては、次の点に注意してください。

    - 最初のパラメーターは、JSON ファイルからのアクションの 1 つです。
    - 2 番目のパラメーターは、JSON ファイル内のアクションにマップされているキーの組み合わせをユーザーが押すと実行される関数です。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. この例を続行するには、最初 `'SHOWTASKPANE'` のパラメーターとして使用します。
1. 関数の本文の場合は [、Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) メソッドを使用してアドインの作業ウィンドウを開きます。 完了すると、コードは次のようになります。

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. 2 番目の関数呼び出しを追加して、アクションを `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide を呼び出す関数にマップします](/javascript/api/office/office.addin#hide--)。 例を次に示します。

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

前の手順に従って、アドインでは **、Ctrl + Shift +** 上方向キーと Ctrl + Shift + 下方向キーを押して、作業ウィンドウの表示/非表示 **を切り替えます**。 これは、サンプルの Excel キーボード ショートカット アドインと [同じ動作です](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。

## <a name="details-and-restrictions"></a>詳細と制限

### <a name="constructing-the-action-objects"></a>アクション オブジェクトの作成

次のガイドラインに従って、オブジェクトの配列内のオブジェクトを指定shortcuts.js`action` します。

- プロパティ名 `id` は `name` 必須です。
- この `id` プロパティは、キーボード ショートカットを使用して呼び出すアクションを一意に識別するために使用されます。
- この `name` プロパティは、アクションを説明するユーザー フレンドリーな文字列である必要があります。 文字 A ~ Z、a ~ z、0 ~ 9、および区切り記号 "-"、"_"、および "+" を組み合わせて指定する必要があります。
- プロパティは省略可能です。 現在サポート `ExecuteFunction` されているのは型のみです。

例を次に示します。

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

### <a name="constructing-the-shortcut-objects"></a>ショートカット オブジェクトの作成

次のガイドラインに従って、オブジェクトの配列内のオブジェクトを指定shortcuts.js`shortcuts` します。

- プロパティ名 `action` 、 `key` および `default` 必須です。
- プロパティの値 `action` は文字列であり、アクション オブジェクトのプロパティの 1 `id` つと一致する必要があります。
- この `default` プロパティには、文字 A ~ Z、-z、0 ~ 9、句読点 "-"、"_"、および "+" を任意に組み合わせて指定できます。 (規則により、これらのプロパティでは小文字は使用されません)。
- この `default` プロパティには、少なくとも 1 つの修飾キー (Alt、Ctrl、Shift) の名前と、他の 1 つのキーのみを含む必要があります。
- Mac では、COMMAND 修飾子キーもサポートされています。
- Mac の場合、ALT キーは OPTION キーにマップされます。 Windows では、COMMAND は Ctrl キーにマップされます。
- 標準キーボードで 2 文字が同じ物理キーにリンクされている場合、それらはプロパティの同義語になります。たとえば、Alt + a と Alt + A は同じショートカットであり、"-" と `default` "_" は同じ物理キーなので、Ctrl + + と Ctrl+ も同じです。 \_
- "+" 文字は、いずれかの側のキーが同時に押された状態を示します。

例を次に示します。

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!NOTE]
> キーヒント (塗りつぶしの色を Alt **+ H、H)** を選択する Excel ショートカットなど、シーケンシャル キー ショートカットとも呼ばれる) は、Office アドインではサポートされていません。

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>作業ウィンドウにフォーカスがあるときにショートカットを使用する

現在、アドインのキーボード ショートカットOffice、ユーザーのフォーカスがワークシートにある場合にのみ呼び出すことができます。 ユーザーのフォーカスが Office UI (作業ウィンドウなど) 内にある場合、アドインのショートカットは無視されません。 回避策として、アドインは、ユーザーのフォーカスがアドイン UI 内にあるときに特定のアクションを呼び出すキーボード ハンドラーを定義できます。

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>他のアドインで既に使用されているOffice組み合わせを使用する

プレビュー期間中に、ユーザーがアドインによって登録されたキーの組み合わせ、および Office または別のアドインによって押された場合の処理を判断するシステムはありません。 動作は未定義です。

現在、2 つ以上のアドインが同じキーボード ショートカットを登録している場合は回避策はありません。ただし、次の優れたプラクティスを使用して Excel との競合を最小限に抑える可能性があります。

- アドインでは、次のパターンのキーボード ショートカットのみを使用します。**Ctrl + Shift + Alt +* x*** *(x* は他のキーです)。
- 追加のキーボード ショートカットが必要な場合は [、Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)のキーボード ショートカットの一覧を確認し、アドインで使用しないようにします。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>上書きできないブラウザー ショートカット

次のキーボードの組み合わせは使用できません。 ブラウザーで使用され、上書きできません。 この一覧は、進行中の作業です。 上書きできない他の組み合わせが見つかった場合は、このページの下部にあるフィードバック ツールを使用してお知らせください。

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + PgUp/PgDn

## <a name="next-steps"></a>次の手順

- サンプル アドインの [excel-keyboard-shortcuts を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。
