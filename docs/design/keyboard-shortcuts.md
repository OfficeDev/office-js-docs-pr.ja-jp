---
title: Office アドインでのユーザー設定のキーボードショートカット
description: Office アドインにキーの組み合わせとも呼ばれるユーザー設定のキーボードショートカットを追加する方法について説明します。
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: 40009dd92787b7c220bb8cfc741cffb2e4b68a9e
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132040"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>カスタムキーボードショートカットを Office アドインに追加する (プレビュー)

キーの組み合わせとも呼ばれるキーボードショートカットを使用すると、アドインのユーザーの作業効率を高めることができます。また、障害が発生したユーザーのためにアドインのアクセシビリティを向上させるために、マウスに代わる手段を提供します。

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> ショートカットキーが有効になっているアドインの作業バージョンから始めるには、サンプルの [Excel キーボードショートカット](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)を複製して実行します。 独自のアドインにキーボードショートカットを追加する準備ができたら、この記事に進みます。

アドインにキーボードショートカットを追加するには、次の3つの手順を実行します。

1. [アドインのマニフェストを構成](#configure-the-manifest)します。
1. [[ショートカット] JSON ファイルを作成または編集](#create-or-edit-the-shortcuts-json-file)して、アクションとそのキーボードショートカットを定義します。
1. 各アクションに関数を[割り当てる API の](/javascript/api/office/office.actions#associate)1 つ以上の[ランタイム呼び出しを追加](#create-a-mapping-of-actions-to-their-functions)します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストに対して2つの小さな変更が行われます。 1つは、アドインで共有ランタイムを使用できるようにし、もう1つは、キーボードショートカットを定義した JSON 形式のファイルを参照することです。

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

カスタムキーボードショートカットを追加するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

### <a name="link-the-mapping-file-to-the-manifest"></a>マッピングファイルをマニフェストにリンクする

マニフェスト内の要素のすぐ *下* に `<VersionOverrides>` 、 [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 要素を追加します (内部は含まれていません)。 この属性を、 `Url` 後の手順で作成するプロジェクト内の JSON ファイルの完全な URL に設定します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>ショートカット JSON ファイルを作成または編集する

プロジェクトに JSON ファイルを作成します。 ファイルのパスが、 `Url` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 要素の属性に指定した場所と一致していることを確認してください。 このファイルは、キーボードショートカットと、それが呼び出すアクションについて説明します。

1. JSON ファイルの内部には、2つの配列があります。 Actions 配列には、呼び出されるアクションを定義するオブジェクトが格納されます。ショートカット配列には、アクションに対するキーの組み合わせをマップするオブジェクトが格納されます。 次に例を示します：

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

    JSON オブジェクトの詳細については、「 [action オブジェクトを構築](#constructing-the-action-objects) する」と「 [ショートカットオブジェクトを構築](#constructing-the-shortcut-objects)する」を参照してください。 JSON の完全なスキーマは [extended-manifest.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)あります。 (メモ: スキーマへのリンクは、プレビュー期間の初期段階では機能しない可能性があります。)

    > [!NOTE]
    > この記事では、"CTRL" の代わりに "CONTROL" を使用できます。

    後の手順では、操作は自分で記述した関数にマップされます。 この例では、メソッドを呼び出す関数に対して、SHOWTASKPANE をこのメソッドを呼び出す関数に対して後でマップし `Office.addin.showAsTaskpane` `Office.addin.hide` ます。

## <a name="create-a-mapping-of-actions-to-their-functions"></a>各機能にアクションのマッピングを作成する

1. プロジェクトで、HTML ページに読み込まれた JavaScript ファイルを要素に開き `<FunctionFile>` ます。
1. JavaScript ファイルで、JSON ファイルで指定した各アクションを JavaScript 関数にマップするのには、「 [Office. actions.](/javascript/api/office/office.actions#associate) 」という関連付け API を使用します。 次の JavaScript をファイルに追加します。 コードについては、次の点に注意してください。

    - 最初のパラメーターは、JSON ファイルからのアクションの1つです。
    - 2番目のパラメーターは、ユーザーが JSON ファイルのアクションにマップされたキーの組み合わせを押したときに実行される関数です。

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. 例を続行するには、 `'SHOWTASKPANE'` 最初のパラメーターとしてを使用します。
1. 関数の本文については、 [Office](/javascript/api/office/office.addin#showastaskpane--) を使用してアドインの作業ウィンドウを開きます。 完了すると、コードは次のようになります。

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

1. 関数の2番目の呼び出しを追加し `Office.actions.associate` `HIDETASKPANE` て、アクションを呼び出す[Office.addin.hide](/javascript/api/office/office.addin#hide--)関数にアクションをマップします。 例を次に示します。

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

前の手順に従って、 **ctrl + shift + 上方向キー** と **ctrl + shift + ↓キー** を押して、アドインで作業ウィンドウの表示を切り替えることができます。 これは、「 [excel キーボードショートカットアドインのサンプル](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)」に記載されているのと同じ動作になります。

## <a name="details-and-restrictions"></a>詳細と制限事項

### <a name="constructing-the-action-objects"></a>Action オブジェクトを構築する

shortcuts.jsの配列内のオブジェクトを指定するときは、次のガイドラインを使用し `action` ます。

- プロパティ名は `id` `name` 必須です。
- この `id` プロパティは、キーボードショートカットを使用して呼び出すアクションを一意に識別するために使用されます。
- この `name` プロパティは、アクションを説明するユーザーフレンドリ文字列である必要があります。 この文字列は、A ~ Z、a ~ z、0-9、および句読点 "-"、"_"、および "+" の文字の組み合わせである必要があります。
- プロパティは省略可能です。 現在 `ExecuteFunction` 、型のみがサポートされています。

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

JSON の完全なスキーマは [extended-manifest.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)あります。 (メモ: スキーマへのリンクは、プレビュー期間の初期段階では機能しない可能性があります。)

### <a name="constructing-the-shortcut-objects"></a>ショートカットオブジェクトを構築する

shortcuts.jsの配列内のオブジェクトを指定するときは、次のガイドラインを使用し `shortcuts` ます。

- プロパティ名、 `action` `key` 、および `default` が必要です。
- プロパティの値 `action` は文字列で、action オブジェクトのプロパティのいずれかに一致している必要があり `id` ます。
- このプロパティには、 `default` a ~ z、a ~ z、0-9、および句読点 "-"、"_"、および "+" の文字を任意に組み合わせて使用できます。 (慣例では、これらのプロパティに小文字は使用されません)。
- このプロパティには、 `default` 少なくとも1つの修飾子キー (ALT、CTRL、SHIFT) の名前と、その他の1つのキーのみを含める必要があります。
- Mac では、コマンド修飾子キーもサポートしています。
- Mac の場合、ALT はオプションキーにマップされます。 Windows の場合、COMMAND は CTRL キーにマップされます。
- 標準キーボードで2つの文字が同じ物理キーにリンクされている場合は、プロパティの類義語です `default` 。たとえば、alt + a と alt + a は同じショートカットです。たとえば、ctrl +-と ctrl + + は同じ \_ 物理キーです。
- "+" 文字は、その両側のキーが同時に押されたことを示します。

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

JSON の完全なスキーマは [extended-manifest.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)あります。 (メモ: スキーマへのリンクは、プレビュー期間の初期段階では機能しない可能性があります。)

> [!NOTE]
> キーヒント (連続したキーショートカットとも呼ばれます)。これは、Office アドインでは、塗りつぶしの色として **Alt + h** を選択するための Excel ショートカットです。

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>作業ウィンドウにフォーカスがあるときにショートカットを使用する

現時点では、Office アドインのキーボードショートカットは、ユーザーのフォーカスがワークシートにある場合にのみ呼び出すことができます。 ユーザーのフォーカスが Office UI (作業ウィンドウなど) 内にある場合、アドインのショートカットは無視されません。 回避策として、アドインでは、ユーザーのフォーカスがアドインの UI 内にあるときに特定のアクションを呼び出すことができるキーボードハンドラーを定義できます。

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Office または他のアドインで既に使用されているキーの組み合わせの使用

プレビュー期間中は、アドインによって登録されたキーの組み合わせと、Office または別のアドインによって登録されたキーの組み合わせをユーザーが押したときに発生する処理を判断するためのシステムはありません。 動作は未定義です。

現時点では、2つ以上のアドインによって同じキーボードショートカットが登録されていても、次のような正しい方法で Excel との競合を最小限に抑えることができます。

- アドインでは次のパターンのキーボードショートカットのみを使用します: **Ctrl + Shift + Alt +* x * * *。 *x* は他のキーです。
- さらに多くのキーボードショートカットが必要な場合は、 [Excel キーボードショートカットの一覧](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)をチェックして、アドインでそのショートカットを使用しないようにします。

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>上書きできないブラウザーショートカット

次のキーの組み合わせは使用できません。 これらはブラウザーで使用され、上書きすることはできません。 このリストは、進行中の作業を示しています。 上書きできない他の組み合わせが見つかった場合は、このページの下部にあるフィードバックツールを使用してお知らせください。

- Ctrl + N
- Ctrl + Shift + N
- Ctrl + T
- Ctrl + Shift + T
- Ctrl + W
- Ctrl + PgUp/PgDn

## <a name="next-steps"></a>次の手順

- サンプルアドインの [excel ショートカットキー](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)を参照してください。
