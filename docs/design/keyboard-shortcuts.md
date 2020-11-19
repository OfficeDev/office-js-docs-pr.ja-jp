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
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="899b2-103">カスタムキーボードショートカットを Office アドインに追加する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="899b2-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="899b2-104">キーの組み合わせとも呼ばれるキーボードショートカットを使用すると、アドインのユーザーの作業効率を高めることができます。また、障害が発生したユーザーのためにアドインのアクセシビリティを向上させるために、マウスに代わる手段を提供します。</span><span class="sxs-lookup"><span data-stu-id="899b2-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="899b2-105">ショートカットキーが有効になっているアドインの作業バージョンから始めるには、サンプルの [Excel キーボードショートカット](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)を複製して実行します。</span><span class="sxs-lookup"><span data-stu-id="899b2-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="899b2-106">独自のアドインにキーボードショートカットを追加する準備ができたら、この記事に進みます。</span><span class="sxs-lookup"><span data-stu-id="899b2-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="899b2-107">アドインにキーボードショートカットを追加するには、次の3つの手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="899b2-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="899b2-108">[アドインのマニフェストを構成](#configure-the-manifest)します。</span><span class="sxs-lookup"><span data-stu-id="899b2-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="899b2-109">[[ショートカット] JSON ファイルを作成または編集](#create-or-edit-the-shortcuts-json-file)して、アクションとそのキーボードショートカットを定義します。</span><span class="sxs-lookup"><span data-stu-id="899b2-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="899b2-110">各アクションに関数を[割り当てる API の](/javascript/api/office/office.actions#associate)1 つ以上の[ランタイム呼び出しを追加](#create-a-mapping-of-actions-to-their-functions)します。</span><span class="sxs-lookup"><span data-stu-id="899b2-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="899b2-111">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="899b2-111">Configure the manifest</span></span>

<span data-ttu-id="899b2-112">マニフェストに対して2つの小さな変更が行われます。</span><span class="sxs-lookup"><span data-stu-id="899b2-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="899b2-113">1つは、アドインで共有ランタイムを使用できるようにし、もう1つは、キーボードショートカットを定義した JSON 形式のファイルを参照することです。</span><span class="sxs-lookup"><span data-stu-id="899b2-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="899b2-114">共有ランタイムを使用するようにアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="899b2-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="899b2-115">カスタムキーボードショートカットを追加するには、アドインで共有ランタイムを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="899b2-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="899b2-116">詳細については、「 [共有ランタイムを使用するようにアドインを構成する](../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="899b2-116">For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="899b2-117">マッピングファイルをマニフェストにリンクする</span><span class="sxs-lookup"><span data-stu-id="899b2-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="899b2-118">マニフェスト内の要素のすぐ *下* に `<VersionOverrides>` 、 [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 要素を追加します (内部は含まれていません)。</span><span class="sxs-lookup"><span data-stu-id="899b2-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="899b2-119">この属性を、 `Url` 後の手順で作成するプロジェクト内の JSON ファイルの完全な URL に設定します。</span><span class="sxs-lookup"><span data-stu-id="899b2-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="899b2-120">ショートカット JSON ファイルを作成または編集する</span><span class="sxs-lookup"><span data-stu-id="899b2-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="899b2-121">プロジェクトに JSON ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="899b2-121">Create a JSON file in your project.</span></span> <span data-ttu-id="899b2-122">ファイルのパスが、 `Url` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 要素の属性に指定した場所と一致していることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="899b2-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="899b2-123">このファイルは、キーボードショートカットと、それが呼び出すアクションについて説明します。</span><span class="sxs-lookup"><span data-stu-id="899b2-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="899b2-124">JSON ファイルの内部には、2つの配列があります。</span><span class="sxs-lookup"><span data-stu-id="899b2-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="899b2-125">Actions 配列には、呼び出されるアクションを定義するオブジェクトが格納されます。ショートカット配列には、アクションに対するキーの組み合わせをマップするオブジェクトが格納されます。</span><span class="sxs-lookup"><span data-stu-id="899b2-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="899b2-126">次に例を示します：</span><span class="sxs-lookup"><span data-stu-id="899b2-126">Here is an example:</span></span>

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

    <span data-ttu-id="899b2-127">JSON オブジェクトの詳細については、「 [action オブジェクトを構築](#constructing-the-action-objects) する」と「 [ショートカットオブジェクトを構築](#constructing-the-shortcut-objects)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="899b2-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="899b2-128">JSON の完全なスキーマは [extended-manifest.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)あります。</span><span class="sxs-lookup"><span data-stu-id="899b2-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="899b2-129">(メモ: スキーマへのリンクは、プレビュー期間の初期段階では機能しない可能性があります。)</span><span class="sxs-lookup"><span data-stu-id="899b2-129">(Note: The link to the schema may not be working early in the preview period.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="899b2-130">この記事では、"CTRL" の代わりに "CONTROL" を使用できます。</span><span class="sxs-lookup"><span data-stu-id="899b2-130">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="899b2-131">後の手順では、操作は自分で記述した関数にマップされます。</span><span class="sxs-lookup"><span data-stu-id="899b2-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="899b2-132">この例では、メソッドを呼び出す関数に対して、SHOWTASKPANE をこのメソッドを呼び出す関数に対して後でマップし `Office.addin.showAsTaskpane` `Office.addin.hide` ます。</span><span class="sxs-lookup"><span data-stu-id="899b2-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="899b2-133">各機能にアクションのマッピングを作成する</span><span class="sxs-lookup"><span data-stu-id="899b2-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="899b2-134">プロジェクトで、HTML ページに読み込まれた JavaScript ファイルを要素に開き `<FunctionFile>` ます。</span><span class="sxs-lookup"><span data-stu-id="899b2-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="899b2-135">JavaScript ファイルで、JSON ファイルで指定した各アクションを JavaScript 関数にマップするのには、「 [Office. actions.](/javascript/api/office/office.actions#associate) 」という関連付け API を使用します。</span><span class="sxs-lookup"><span data-stu-id="899b2-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="899b2-136">次の JavaScript をファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="899b2-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="899b2-137">コードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="899b2-137">Note the following about the code:</span></span>

    - <span data-ttu-id="899b2-138">最初のパラメーターは、JSON ファイルからのアクションの1つです。</span><span class="sxs-lookup"><span data-stu-id="899b2-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="899b2-139">2番目のパラメーターは、ユーザーが JSON ファイルのアクションにマップされたキーの組み合わせを押したときに実行される関数です。</span><span class="sxs-lookup"><span data-stu-id="899b2-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="899b2-140">例を続行するには、 `'SHOWTASKPANE'` 最初のパラメーターとしてを使用します。</span><span class="sxs-lookup"><span data-stu-id="899b2-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="899b2-141">関数の本文については、 [Office](/javascript/api/office/office.addin#showastaskpane--) を使用してアドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="899b2-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="899b2-142">完了すると、コードは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="899b2-142">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="899b2-143">関数の2番目の呼び出しを追加し `Office.actions.associate` `HIDETASKPANE` て、アクションを呼び出す[Office.addin.hide](/javascript/api/office/office.addin#hide--)関数にアクションをマップします。</span><span class="sxs-lookup"><span data-stu-id="899b2-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="899b2-144">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="899b2-144">The following is an example:</span></span>

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

<span data-ttu-id="899b2-145">前の手順に従って、 **ctrl + shift + 上方向キー** と **ctrl + shift + ↓キー** を押して、アドインで作業ウィンドウの表示を切り替えることができます。</span><span class="sxs-lookup"><span data-stu-id="899b2-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="899b2-146">これは、「 [excel キーボードショートカットアドインのサンプル](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)」に記載されているのと同じ動作になります。</span><span class="sxs-lookup"><span data-stu-id="899b2-146">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="899b2-147">詳細と制限事項</span><span class="sxs-lookup"><span data-stu-id="899b2-147">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="899b2-148">Action オブジェクトを構築する</span><span class="sxs-lookup"><span data-stu-id="899b2-148">Constructing the action objects</span></span>

<span data-ttu-id="899b2-149">shortcuts.jsの配列内のオブジェクトを指定するときは、次のガイドラインを使用し `action` ます。</span><span class="sxs-lookup"><span data-stu-id="899b2-149">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="899b2-150">プロパティ名は `id` `name` 必須です。</span><span class="sxs-lookup"><span data-stu-id="899b2-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="899b2-151">この `id` プロパティは、キーボードショートカットを使用して呼び出すアクションを一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="899b2-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="899b2-152">この `name` プロパティは、アクションを説明するユーザーフレンドリ文字列である必要があります。</span><span class="sxs-lookup"><span data-stu-id="899b2-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="899b2-153">この文字列は、A ~ Z、a ~ z、0-9、および句読点 "-"、"_"、および "+" の文字の組み合わせである必要があります。</span><span class="sxs-lookup"><span data-stu-id="899b2-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="899b2-154">プロパティは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="899b2-154">The `type` property is optional.</span></span> <span data-ttu-id="899b2-155">現在 `ExecuteFunction` 、型のみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="899b2-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="899b2-156">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="899b2-156">The following is an example:</span></span>

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

<span data-ttu-id="899b2-157">JSON の完全なスキーマは [extended-manifest.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)あります。</span><span class="sxs-lookup"><span data-stu-id="899b2-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="899b2-158">(メモ: スキーマへのリンクは、プレビュー期間の初期段階では機能しない可能性があります。)</span><span class="sxs-lookup"><span data-stu-id="899b2-158">(Note: The link to the schema may not be working early in the preview period.)</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="899b2-159">ショートカットオブジェクトを構築する</span><span class="sxs-lookup"><span data-stu-id="899b2-159">Constructing the shortcut objects</span></span>

<span data-ttu-id="899b2-160">shortcuts.jsの配列内のオブジェクトを指定するときは、次のガイドラインを使用し `shortcuts` ます。</span><span class="sxs-lookup"><span data-stu-id="899b2-160">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="899b2-161">プロパティ名、 `action` `key` 、および `default` が必要です。</span><span class="sxs-lookup"><span data-stu-id="899b2-161">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="899b2-162">プロパティの値 `action` は文字列で、action オブジェクトのプロパティのいずれかに一致している必要があり `id` ます。</span><span class="sxs-lookup"><span data-stu-id="899b2-162">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="899b2-163">このプロパティには、 `default` a ~ z、a ~ z、0-9、および句読点 "-"、"_"、および "+" の文字を任意に組み合わせて使用できます。</span><span class="sxs-lookup"><span data-stu-id="899b2-163">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="899b2-164">(慣例では、これらのプロパティに小文字は使用されません)。</span><span class="sxs-lookup"><span data-stu-id="899b2-164">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="899b2-165">このプロパティには、 `default` 少なくとも1つの修飾子キー (ALT、CTRL、SHIFT) の名前と、その他の1つのキーのみを含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="899b2-165">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="899b2-166">Mac では、コマンド修飾子キーもサポートしています。</span><span class="sxs-lookup"><span data-stu-id="899b2-166">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="899b2-167">Mac の場合、ALT はオプションキーにマップされます。</span><span class="sxs-lookup"><span data-stu-id="899b2-167">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="899b2-168">Windows の場合、COMMAND は CTRL キーにマップされます。</span><span class="sxs-lookup"><span data-stu-id="899b2-168">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="899b2-169">標準キーボードで2つの文字が同じ物理キーにリンクされている場合は、プロパティの類義語です `default` 。たとえば、alt + a と alt + a は同じショートカットです。たとえば、ctrl +-と ctrl + + は同じ \_ 物理キーです。</span><span class="sxs-lookup"><span data-stu-id="899b2-169">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="899b2-170">"+" 文字は、その両側のキーが同時に押されたことを示します。</span><span class="sxs-lookup"><span data-stu-id="899b2-170">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="899b2-171">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="899b2-171">The following is an example:</span></span>

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

<span data-ttu-id="899b2-172">JSON の完全なスキーマは [extended-manifest.schema.jsに](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)あります。</span><span class="sxs-lookup"><span data-stu-id="899b2-172">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span> <span data-ttu-id="899b2-173">(メモ: スキーマへのリンクは、プレビュー期間の初期段階では機能しない可能性があります。)</span><span class="sxs-lookup"><span data-stu-id="899b2-173">(Note: The link to the schema may not be working early in the preview period.)</span></span>

> [!NOTE]
> <span data-ttu-id="899b2-174">キーヒント (連続したキーショートカットとも呼ばれます)。これは、Office アドインでは、塗りつぶしの色として **Alt + h** を選択するための Excel ショートカットです。</span><span class="sxs-lookup"><span data-stu-id="899b2-174">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="899b2-175">作業ウィンドウにフォーカスがあるときにショートカットを使用する</span><span class="sxs-lookup"><span data-stu-id="899b2-175">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="899b2-176">現時点では、Office アドインのキーボードショートカットは、ユーザーのフォーカスがワークシートにある場合にのみ呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="899b2-176">Currently, the keyboard shortcuts for an Office add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="899b2-177">ユーザーのフォーカスが Office UI (作業ウィンドウなど) 内にある場合、アドインのショートカットは無視されません。</span><span class="sxs-lookup"><span data-stu-id="899b2-177">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="899b2-178">回避策として、アドインでは、ユーザーのフォーカスがアドインの UI 内にあるときに特定のアクションを呼び出すことができるキーボードハンドラーを定義できます。</span><span class="sxs-lookup"><span data-stu-id="899b2-178">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="899b2-179">Office または他のアドインで既に使用されているキーの組み合わせの使用</span><span class="sxs-lookup"><span data-stu-id="899b2-179">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="899b2-180">プレビュー期間中は、アドインによって登録されたキーの組み合わせと、Office または別のアドインによって登録されたキーの組み合わせをユーザーが押したときに発生する処理を判断するためのシステムはありません。</span><span class="sxs-lookup"><span data-stu-id="899b2-180">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="899b2-181">動作は未定義です。</span><span class="sxs-lookup"><span data-stu-id="899b2-181">Behavior is undefined.</span></span>

<span data-ttu-id="899b2-182">現時点では、2つ以上のアドインによって同じキーボードショートカットが登録されていても、次のような正しい方法で Excel との競合を最小限に抑えることができます。</span><span class="sxs-lookup"><span data-stu-id="899b2-182">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="899b2-183">アドインでは次のパターンのキーボードショートカットのみを使用します: \**Ctrl + Shift + Alt +* x \* \* \*。 *x* は他のキーです。</span><span class="sxs-lookup"><span data-stu-id="899b2-183">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="899b2-184">さらに多くのキーボードショートカットが必要な場合は、 [Excel キーボードショートカットの一覧](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)をチェックして、アドインでそのショートカットを使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="899b2-184">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="899b2-185">上書きできないブラウザーショートカット</span><span class="sxs-lookup"><span data-stu-id="899b2-185">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="899b2-186">次のキーの組み合わせは使用できません。</span><span class="sxs-lookup"><span data-stu-id="899b2-186">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="899b2-187">これらはブラウザーで使用され、上書きすることはできません。</span><span class="sxs-lookup"><span data-stu-id="899b2-187">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="899b2-188">このリストは、進行中の作業を示しています。</span><span class="sxs-lookup"><span data-stu-id="899b2-188">This list is a work in progress.</span></span> <span data-ttu-id="899b2-189">上書きできない他の組み合わせが見つかった場合は、このページの下部にあるフィードバックツールを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="899b2-189">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="899b2-190">Ctrl + N</span><span class="sxs-lookup"><span data-stu-id="899b2-190">Ctrl+N</span></span>
- <span data-ttu-id="899b2-191">Ctrl + Shift + N</span><span class="sxs-lookup"><span data-stu-id="899b2-191">Ctrl+Shift+N</span></span>
- <span data-ttu-id="899b2-192">Ctrl + T</span><span class="sxs-lookup"><span data-stu-id="899b2-192">Ctrl+T</span></span>
- <span data-ttu-id="899b2-193">Ctrl + Shift + T</span><span class="sxs-lookup"><span data-stu-id="899b2-193">Ctrl+Shift+T</span></span>
- <span data-ttu-id="899b2-194">Ctrl + W</span><span class="sxs-lookup"><span data-stu-id="899b2-194">Ctrl+W</span></span>
- <span data-ttu-id="899b2-195">Ctrl + PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="899b2-195">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="899b2-196">次の手順</span><span class="sxs-lookup"><span data-stu-id="899b2-196">Next Steps</span></span>

- <span data-ttu-id="899b2-197">サンプルアドインの [excel ショートカットキー](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="899b2-197">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
