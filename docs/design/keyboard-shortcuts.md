---
title: カスタム キーボード ショートカット (Office アドイン)
description: カスタム キーボード ショートカット (キーの組み合わせとも呼ばれる) をアドインに追加するOffice説明します。
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: c767c6d5bc23f0a44422452839cd8bdf87bd8715
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505200"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>カスタム キーボード ショートカットをアドインにOfficeする (プレビュー)

キーボード ショートカット (キーの組み合わせとも呼ばれる) を使用すると、アドインのユーザーの作業効率が向上し、マウスの代替手段を提供することで、障がいのあるユーザーに対するアドインのアクセシビリティが向上します。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> キーボード ショートカットが既に有効になっているアドインの作業バージョンから開始するには、 [サンプルの Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)キーボード ショートカットを複製して実行します。 キーボード ショートカットを独自のアドインに追加する準備ができたら、この記事に進む。

アドインにキーボード ショートカットを追加するには、次の 3 つの手順があります。

1. [アドインのマニフェストを構成します](#configure-the-manifest)。
1. [アクションとそのキーボード ショートカットを](#create-or-edit-the-shortcuts-json-file) 定義するショートカット JSON ファイルを作成または編集します。
1. [](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate API の 1 つ以上の](/javascript/api/office/office.actions#associate)ランタイム呼び出しを追加して、関数を各アクションにマップします。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストには 2 つの小さな変更があります。 1 つは、共有ランタイムを使用するアドインを有効にし、もう 1 つは、キーボード ショートカットを定義した JSON 形式のファイルをポイントすることです。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

カスタム キーボード ショートカットを追加するには、共有ランタイムを使用するアドインが必要です。 詳細については、「 [共有ランタイムを使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

### <a name="link-the-mapping-file-to-the-manifest"></a>マッピング ファイルをマニフェストにリンクする

マニフェスト *内* の要素の直下 (内部ではない) `<VersionOverrides>` に [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。 後の手順で作成するプロジェクトの JSON ファイルの完全な URL に属性 `Url` を設定します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>ショートカット JSON ファイルを作成または編集する

プロジェクトに JSON ファイルを作成します。 ファイルのパスが ExtendedOverrides 要素の属性に指定した場所と一致 `Url` [する必要](../reference/manifest/extendedoverrides.md) があります。 このファイルには、キーボード ショートカットと、キーボード ショートカットが呼び出すアクションが記述されます。

1. JSON ファイル内には、2 つの配列があります。 actions 配列には、呼び出すアクションを定義するオブジェクトが含まれます。ショートカット配列には、キーの組み合わせをアクションにマップするオブジェクトが含まれます。 次に例を示します：

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

    JSON オブジェクトの詳細については、「アクション[](#constructing-the-action-objects)オブジェクトの作成」および「ショートカット オブジェクトの[作成」を参照してください](#constructing-the-shortcut-objects)。 ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

    > [!NOTE]
    > この記事では、"CTRL" の代りで "CONTROL" を使用できます。

    後の手順では、アクション自体が作成する関数にマップされます。 この例では、後で SHOWTASKPANE をメソッドを呼び出す関数にマップし、HIDETASKPANE をメソッドを呼び出す `Office.addin.showAsTaskpane` 関数にマップ `Office.addin.hide` します。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>アクションの関数へのマッピングを作成する

1. プロジェクトで、HTML ページによって読み込まれた JavaScript ファイルを要素で開 `<FunctionFile>` きます。
1. JavaScript ファイルで [、Office.actions.associate](/javascript/api/office/office.actions#associate) API を使用して、JSON ファイルで指定した各アクションを JavaScript 関数にマップします。 次の JavaScript をファイルに追加します。 コードについて次の点に注意してください。

    - 最初のパラメーターは、JSON ファイルからのアクションの 1 つです。
    - 2 番目のパラメーターは、JSON ファイル内のアクションにマップされているキーの組み合わせをユーザーが押すと実行される関数です。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. この例を続行するには、最初 `'SHOWTASKPANE'` のパラメーターとして使用します。
1. 関数の本文では [、Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) メソッドを使用してアドインの作業ウィンドウを開きます。 完了したら、コードは次のようになります。

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

1. 2 つ目の関数呼び出しを追加して、アクション `Office.actions.associate` `HIDETASKPANE` を [Office.addin.hide を](/javascript/api/office/office.addin#hide--)呼び出す関数にマップします。 例を次に示します。

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

前の手順に従うと **、Ctrl + Shift + 上** 矢印キーと Ctrl + Shift + 下矢印キーを押して、作業ウィンドウの表示を切り **替えます**。 これは、サンプルの Excel キーボード ショートカット アドインに示されている動作 [と同じです](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。

## <a name="details-and-restrictions"></a>詳細と制限

### <a name="constructing-the-action-objects"></a>アクション オブジェクトの作成

次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`action` します。

- プロパティ名 `id` と `name` 必須です。
- この `id` プロパティは、キーボード ショートカットを使用して呼び出すアクションを一意に識別するために使用されます。
- プロパティ `name` は、アクションを記述するユーザーフレンドリーな文字列である必要があります。 文字 A - Z、a - z、0 ~ 9、および句読点 "-"、"_"、および "+" の組み合わせである必要があります。
- プロパティは省略可能です。 現在は `ExecuteFunction` 型のみサポートされています。

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

次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`shortcuts` します。

- プロパティ名 `action` 、 `key` および `default` 必須です。
- プロパティの値は `action` 文字列であり、action オブジェクトのプロパティの 1 `id` つと一致する必要があります。
- プロパティ `default` には、文字 A ~ Z、-z、0 ~ 9、句読点 "-"、"_"、"+" の任意の組み合わせを指定できます。 (慣例では、これらのプロパティでは小文字は使用されません)。
- プロパティ `default` には、少なくとも 1 つの修飾子キー (ALT、Ctrl、SHIFT) の名前と、他の 1 つのキーのみを含む必要があります。
- Mac では、COMMAND 修飾子キーもサポートしています。
- Mac の場合、ALT は OPTION キーにマップされます。 Windows の場合、COMMAND は Ctrl キーにマップされます。
- 標準キーボードで 2 つの文字が同じ物理キーにリンクされている場合は、プロパティ内の類義語になります。たとえば、ALT +a と ALT+A は同じショートカットなので、"-" と "_" は同じ物理キーなので `default` 、Ctrl ++ と Ctrl+ も同様です。 \_
- "+" 文字は、そのいずれかの側のキーが同時に押された状態を示します。

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
> キーヒントは、塗りつぶしの色 **Alt +H、H** を選択する Excel ショートカットなどのシーケンシャル キー ショートカットとも呼ばれる、Office アドインではサポートされていません。

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>作業ウィンドウにフォーカスがあるときにショートカットを使用する

現在、ユーザーのフォーカスがワークシートにある場合Officeアドインのキーボード ショートカットを呼び出すことができます。 ユーザーのフォーカスが作業ウィンドウOffice UI 内にある場合、アドインのショートカットは無視されません。 回避策として、アドインは、ユーザーのフォーカスがアドイン UI 内にあるときに特定のアクションを呼び出すキーボード ハンドラーを定義できます。

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>ユーザーまたは別のアドインで既にOfficeキーの組み合わせを使用する

プレビュー期間中、ユーザーがアドインによって登録されているキーの組み合わせを押すと、Office または別のアドインによって何が起こるかを判断するシステムはありません。 動作は未定義です。

現在、2 つ以上のアドインが同じキーボード ショートカットを登録している場合、回避策はありません。ただし、Excel との競合を最小限に抑えるために、次の方法を使用できます。

- アドインでは、キーボード ショートカットのみを使用します *。*Ctrl +Shift+Alt+* x***、x は他のキーです。
- キーボード ショートカットが必要な場合は [、Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)キーボード ショートカットの一覧を確認し、アドインで使用しないようにします。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>オーバーライドできないブラウザー のショートカット

次のキーボードの組み合わせを使用することはできません。 ブラウザーで使用され、オーバーライドすることはできません。 このリストは進行中の作業です。 上書きできない他の組み合わせを発見した場合は、このページの下部にあるフィードバック ツールを使用してお知らせください。

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + PgUp/PgDn

## <a name="localize-the-keyboard-shortcuts-json"></a>キーボード ショートカット JSON をローカライズする

アドインが複数のローカライズをサポートしている場合は、アクション オブジェクトのプロパティをローカライズ `name` する必要があります。 また、アドインがサポートするローカライズの中にアルファベットや異なる書き込みシステムがある場合、キーボードが異なる場合は、ショートカットのローカライズも必要な場合があります。 キーボード ショートカット JSON をローカライズする方法については、「拡張オーバーライドをローカライズする [」を参照してください](../develop/localization.md#localize-extended-overrides)。

## <a name="next-steps"></a>次の手順

- サンプル アドインの [excel-keyboard-shortcuts を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。
- 「マニフェストの拡張オーバーライドを処理する」の拡張オーバーライドの操作 [の概要を取得します](../develop/extended-overrides.md)。
