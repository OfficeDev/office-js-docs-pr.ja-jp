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
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="49ab5-103">カスタム キーボード ショートカットをアドインにOfficeする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="49ab5-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="49ab5-104">キーボード ショートカット (キーの組み合わせとも呼ばれる) を使用すると、アドインのユーザーの作業効率が向上し、マウスの代替手段を提供することで、障がいのあるユーザーに対するアドインのアクセシビリティが向上します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="49ab5-105">キーボード ショートカットが既に有効になっているアドインの作業バージョンから開始するには、 [サンプルの Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)キーボード ショートカットを複製して実行します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="49ab5-106">キーボード ショートカットを独自のアドインに追加する準備ができたら、この記事に進む。</span><span class="sxs-lookup"><span data-stu-id="49ab5-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="49ab5-107">アドインにキーボード ショートカットを追加するには、次の 3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="49ab5-108">[アドインのマニフェストを構成します](#configure-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="49ab5-109">[アクションとそのキーボード ショートカットを](#create-or-edit-the-shortcuts-json-file) 定義するショートカット JSON ファイルを作成または編集します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="49ab5-110">[](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate API の 1 つ以上の](/javascript/api/office/office.actions#associate)ランタイム呼び出しを追加して、関数を各アクションにマップします。</span><span class="sxs-lookup"><span data-stu-id="49ab5-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="49ab5-111">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="49ab5-111">Configure the manifest</span></span>

<span data-ttu-id="49ab5-112">マニフェストには 2 つの小さな変更があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="49ab5-113">1 つは、共有ランタイムを使用するアドインを有効にし、もう 1 つは、キーボード ショートカットを定義した JSON 形式のファイルをポイントすることです。</span><span class="sxs-lookup"><span data-stu-id="49ab5-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="49ab5-114">共有ランタイムを使用するアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="49ab5-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="49ab5-115">カスタム キーボード ショートカットを追加するには、共有ランタイムを使用するアドインが必要です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="49ab5-116">詳細については、「 [共有ランタイムを使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-116">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="49ab5-117">マッピング ファイルをマニフェストにリンクする</span><span class="sxs-lookup"><span data-stu-id="49ab5-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="49ab5-118">マニフェスト *内* の要素の直下 (内部ではない) `<VersionOverrides>` に [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="49ab5-119">後の手順で作成するプロジェクトの JSON ファイルの完全な URL に属性 `Url` を設定します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="49ab5-120">ショートカット JSON ファイルを作成または編集する</span><span class="sxs-lookup"><span data-stu-id="49ab5-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="49ab5-121">プロジェクトに JSON ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-121">Create a JSON file in your project.</span></span> <span data-ttu-id="49ab5-122">ファイルのパスが ExtendedOverrides 要素の属性に指定した場所と一致 `Url` [する必要](../reference/manifest/extendedoverrides.md) があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="49ab5-123">このファイルには、キーボード ショートカットと、キーボード ショートカットが呼び出すアクションが記述されます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="49ab5-124">JSON ファイル内には、2 つの配列があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="49ab5-125">actions 配列には、呼び出すアクションを定義するオブジェクトが含まれます。ショートカット配列には、キーの組み合わせをアクションにマップするオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="49ab5-126">次に例を示します：</span><span class="sxs-lookup"><span data-stu-id="49ab5-126">Here is an example:</span></span>

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

    <span data-ttu-id="49ab5-127">JSON オブジェクトの詳細については、「アクション[](#constructing-the-action-objects)オブジェクトの作成」および「ショートカット オブジェクトの[作成」を参照してください](#constructing-the-shortcut-objects)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="49ab5-128">ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="49ab5-129">この記事では、"CTRL" の代りで "CONTROL" を使用できます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="49ab5-130">後の手順では、アクション自体が作成する関数にマップされます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="49ab5-131">この例では、後で SHOWTASKPANE をメソッドを呼び出す関数にマップし、HIDETASKPANE をメソッドを呼び出す `Office.addin.showAsTaskpane` 関数にマップ `Office.addin.hide` します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="49ab5-132">アクションの関数へのマッピングを作成する</span><span class="sxs-lookup"><span data-stu-id="49ab5-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="49ab5-133">プロジェクトで、HTML ページによって読み込まれた JavaScript ファイルを要素で開 `<FunctionFile>` きます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="49ab5-134">JavaScript ファイルで [、Office.actions.associate](/javascript/api/office/office.actions#associate) API を使用して、JSON ファイルで指定した各アクションを JavaScript 関数にマップします。</span><span class="sxs-lookup"><span data-stu-id="49ab5-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="49ab5-135">次の JavaScript をファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="49ab5-136">コードについて次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="49ab5-136">Note the following about the code:</span></span>

    - <span data-ttu-id="49ab5-137">最初のパラメーターは、JSON ファイルからのアクションの 1 つです。</span><span class="sxs-lookup"><span data-stu-id="49ab5-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="49ab5-138">2 番目のパラメーターは、JSON ファイル内のアクションにマップされているキーの組み合わせをユーザーが押すと実行される関数です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="49ab5-139">この例を続行するには、最初 `'SHOWTASKPANE'` のパラメーターとして使用します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="49ab5-140">関数の本文では [、Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) メソッドを使用してアドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="49ab5-141">完了したら、コードは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="49ab5-142">2 つ目の関数呼び出しを追加して、アクション `Office.actions.associate` `HIDETASKPANE` を [Office.addin.hide を](/javascript/api/office/office.addin#hide--)呼び出す関数にマップします。</span><span class="sxs-lookup"><span data-stu-id="49ab5-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="49ab5-143">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-143">The following is an example:</span></span>

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

<span data-ttu-id="49ab5-144">前の手順に従うと **、Ctrl + Shift + 上** 矢印キーと Ctrl + Shift + 下矢印キーを押して、作業ウィンドウの表示を切り **替えます**。</span><span class="sxs-lookup"><span data-stu-id="49ab5-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="49ab5-145">これは、サンプルの Excel キーボード ショートカット アドインに示されている動作 [と同じです](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="49ab5-146">詳細と制限</span><span class="sxs-lookup"><span data-stu-id="49ab5-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="49ab5-147">アクション オブジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="49ab5-147">Constructing the action objects</span></span>

<span data-ttu-id="49ab5-148">次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`action` します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="49ab5-149">プロパティ名 `id` と `name` 必須です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="49ab5-150">この `id` プロパティは、キーボード ショートカットを使用して呼び出すアクションを一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="49ab5-151">プロパティ `name` は、アクションを記述するユーザーフレンドリーな文字列である必要があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="49ab5-152">文字 A - Z、a - z、0 ~ 9、および句読点 "-"、"_"、および "+" の組み合わせである必要があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="49ab5-153">プロパティは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-153">The `type` property is optional.</span></span> <span data-ttu-id="49ab5-154">現在は `ExecuteFunction` 型のみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="49ab5-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="49ab5-155">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-155">The following is an example:</span></span>

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

<span data-ttu-id="49ab5-156">ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="49ab5-157">ショートカット オブジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="49ab5-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="49ab5-158">次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`shortcuts` します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="49ab5-159">プロパティ名 `action` 、 `key` および `default` 必須です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="49ab5-160">プロパティの値は `action` 文字列であり、action オブジェクトのプロパティの 1 `id` つと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="49ab5-161">プロパティ `default` には、文字 A ~ Z、-z、0 ~ 9、句読点 "-"、"_"、"+" の任意の組み合わせを指定できます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="49ab5-162">(慣例では、これらのプロパティでは小文字は使用されません)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="49ab5-163">プロパティ `default` には、少なくとも 1 つの修飾子キー (ALT、Ctrl、SHIFT) の名前と、他の 1 つのキーのみを含む必要があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="49ab5-164">Mac では、COMMAND 修飾子キーもサポートしています。</span><span class="sxs-lookup"><span data-stu-id="49ab5-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="49ab5-165">Mac の場合、ALT は OPTION キーにマップされます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="49ab5-166">Windows の場合、COMMAND は Ctrl キーにマップされます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="49ab5-167">標準キーボードで 2 つの文字が同じ物理キーにリンクされている場合は、プロパティ内の類義語になります。たとえば、ALT +a と ALT+A は同じショートカットなので、"-" と "_" は同じ物理キーなので `default` 、Ctrl ++ と Ctrl+ も同様です。 \_</span><span class="sxs-lookup"><span data-stu-id="49ab5-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="49ab5-168">"+" 文字は、そのいずれかの側のキーが同時に押された状態を示します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="49ab5-169">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="49ab5-169">The following is an example:</span></span>

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

<span data-ttu-id="49ab5-170">ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="49ab5-171">キーヒントは、塗りつぶしの色 **Alt +H、H** を選択する Excel ショートカットなどのシーケンシャル キー ショートカットとも呼ばれる、Office アドインではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="49ab5-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="49ab5-172">作業ウィンドウにフォーカスがあるときにショートカットを使用する</span><span class="sxs-lookup"><span data-stu-id="49ab5-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="49ab5-173">現在、ユーザーのフォーカスがワークシートにある場合Officeアドインのキーボード ショートカットを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-173">Currently, the keyboard shortcuts for an Office Add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="49ab5-174">ユーザーのフォーカスが作業ウィンドウOffice UI 内にある場合、アドインのショートカットは無視されません。</span><span class="sxs-lookup"><span data-stu-id="49ab5-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="49ab5-175">回避策として、アドインは、ユーザーのフォーカスがアドイン UI 内にあるときに特定のアクションを呼び出すキーボード ハンドラーを定義できます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="49ab5-176">ユーザーまたは別のアドインで既にOfficeキーの組み合わせを使用する</span><span class="sxs-lookup"><span data-stu-id="49ab5-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="49ab5-177">プレビュー期間中、ユーザーがアドインによって登録されているキーの組み合わせを押すと、Office または別のアドインによって何が起こるかを判断するシステムはありません。</span><span class="sxs-lookup"><span data-stu-id="49ab5-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="49ab5-178">動作は未定義です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-178">Behavior is undefined.</span></span>

<span data-ttu-id="49ab5-179">現在、2 つ以上のアドインが同じキーボード ショートカットを登録している場合、回避策はありません。ただし、Excel との競合を最小限に抑えるために、次の方法を使用できます。</span><span class="sxs-lookup"><span data-stu-id="49ab5-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="49ab5-180">アドインでは、キーボード ショートカットのみを使用します *。*Ctrl +Shift+Alt+\* x\*\*\*、x は他のキーです。</span><span class="sxs-lookup"><span data-stu-id="49ab5-180">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="49ab5-181">キーボード ショートカットが必要な場合は [、Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)キーボード ショートカットの一覧を確認し、アドインで使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="49ab5-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="49ab5-182">オーバーライドできないブラウザー のショートカット</span><span class="sxs-lookup"><span data-stu-id="49ab5-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="49ab5-183">次のキーボードの組み合わせを使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="49ab5-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="49ab5-184">ブラウザーで使用され、オーバーライドすることはできません。</span><span class="sxs-lookup"><span data-stu-id="49ab5-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="49ab5-185">このリストは進行中の作業です。</span><span class="sxs-lookup"><span data-stu-id="49ab5-185">This list is a work in progress.</span></span> <span data-ttu-id="49ab5-186">上書きできない他の組み合わせを発見した場合は、このページの下部にあるフィードバック ツールを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="49ab5-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="49ab5-187">Ctrl + N</span><span class="sxs-lookup"><span data-stu-id="49ab5-187">Ctrl+N</span></span>
- <span data-ttu-id="49ab5-188">Ctrl + Shift + N</span><span class="sxs-lookup"><span data-stu-id="49ab5-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="49ab5-189">Ctrl + T</span><span class="sxs-lookup"><span data-stu-id="49ab5-189">Ctrl+T</span></span>
- <span data-ttu-id="49ab5-190">Ctrl + Shift + T</span><span class="sxs-lookup"><span data-stu-id="49ab5-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="49ab5-191">Ctrl + W</span><span class="sxs-lookup"><span data-stu-id="49ab5-191">Ctrl+W</span></span>
- <span data-ttu-id="49ab5-192">Ctrl + PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="49ab5-192">Ctrl+PgUp/PgDn</span></span>

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="49ab5-193">キーボード ショートカット JSON をローカライズする</span><span class="sxs-lookup"><span data-stu-id="49ab5-193">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="49ab5-194">アドインが複数のローカライズをサポートしている場合は、アクション オブジェクトのプロパティをローカライズ `name` する必要があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-194">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="49ab5-195">また、アドインがサポートするローカライズの中にアルファベットや異なる書き込みシステムがある場合、キーボードが異なる場合は、ショートカットのローカライズも必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="49ab5-195">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="49ab5-196">キーボード ショートカット JSON をローカライズする方法については、「拡張オーバーライドをローカライズする [」を参照してください](../develop/localization.md#localize-extended-overrides)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-196">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="next-steps"></a><span data-ttu-id="49ab5-197">次の手順</span><span class="sxs-lookup"><span data-stu-id="49ab5-197">Next Steps</span></span>

- <span data-ttu-id="49ab5-198">サンプル アドインの [excel-keyboard-shortcuts を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-198">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
- <span data-ttu-id="49ab5-199">「マニフェストの拡張オーバーライドを処理する」の拡張オーバーライドの操作 [の概要を取得します](../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="49ab5-199">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
