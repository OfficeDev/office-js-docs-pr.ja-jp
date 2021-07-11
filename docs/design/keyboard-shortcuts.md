---
title: カスタム キーボード ショートカット (Office アドイン)
description: カスタム キーボード ショートカット (キーの組み合わせとも呼ばれる) をアドインに追加するOffice説明します。
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: de8ce0d89dca6745cba96ac9a5ea946d50d41de4
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349257"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a><span data-ttu-id="30289-103">カスタム キーボード ショートカットをアドインOffice追加する</span><span class="sxs-lookup"><span data-stu-id="30289-103">Add custom keyboard shortcuts to your Office Add-ins</span></span>

<span data-ttu-id="30289-104">キーボード ショートカット (キーの組み合わせとも呼ばれる) を使用すると、アドインのユーザーの作業効率が向上します。</span><span class="sxs-lookup"><span data-stu-id="30289-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently.</span></span> <span data-ttu-id="30289-105">キーボード ショートカットは、マウスの代替手段を提供することで、障がいを持つユーザーに対するアドインのアクセシビリティも向上します。</span><span class="sxs-lookup"><span data-stu-id="30289-105">Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="30289-106">キーボード ショートカットが既に有効になっているアドインの作業バージョンから開始するには、キーボード ショートカットのサンプルを複製Excel[実行します](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)。</span><span class="sxs-lookup"><span data-stu-id="30289-106">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="30289-107">キーボード ショートカットを独自のアドインに追加する準備ができたら、この記事に進む。</span><span class="sxs-lookup"><span data-stu-id="30289-107">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="30289-108">アドインにキーボード ショートカットを追加するには、次の 3 つの手順があります。</span><span class="sxs-lookup"><span data-stu-id="30289-108">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="30289-109">[アドインのマニフェストを構成します](#configure-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="30289-109">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="30289-110">[アクションとそのキーボード ショートカットを](#create-or-edit-the-shortcuts-json-file) 定義するショートカット JSON ファイルを作成または編集します。</span><span class="sxs-lookup"><span data-stu-id="30289-110">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="30289-111">[](#create-a-mapping-of-actions-to-their-functions) [Office.actions.associate](/javascript/api/office/office.actions#associate) API の 1 つ以上のランタイム呼び出しを追加して、各アクションに関数をマップします。</span><span class="sxs-lookup"><span data-stu-id="30289-111">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="30289-112">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="30289-112">Configure the manifest</span></span>

<span data-ttu-id="30289-113">マニフェストには 2 つの小さな変更があります。</span><span class="sxs-lookup"><span data-stu-id="30289-113">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="30289-114">1 つは、共有ランタイムを使用するアドインを有効にし、もう 1 つは、キーボード ショートカットを定義した JSON 形式のファイルをポイントすることです。</span><span class="sxs-lookup"><span data-stu-id="30289-114">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="30289-115">共有ランタイムを使用するアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="30289-115">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="30289-116">カスタム キーボード ショートカットを追加するには、共有ランタイムを使用するアドインが必要です。</span><span class="sxs-lookup"><span data-stu-id="30289-116">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="30289-117">詳細については、「 [共有ランタイムを使用するアドインを構成する」を参照してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="30289-117">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="30289-118">マッピング ファイルをマニフェストにリンクする</span><span class="sxs-lookup"><span data-stu-id="30289-118">Link the mapping file to the manifest</span></span>

<span data-ttu-id="30289-119">マニフェスト *内* の要素の直下 (内部ではない) `<VersionOverrides>` に [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。</span><span class="sxs-lookup"><span data-stu-id="30289-119">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="30289-120">後の手順で作成するプロジェクトの JSON ファイルの完全な URL に属性 `Url` を設定します。</span><span class="sxs-lookup"><span data-stu-id="30289-120">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="30289-121">ショートカット JSON ファイルを作成または編集する</span><span class="sxs-lookup"><span data-stu-id="30289-121">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="30289-122">プロジェクトに JSON ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="30289-122">Create a JSON file in your project.</span></span> <span data-ttu-id="30289-123">ファイルのパスが ExtendedOverrides 要素の属性に指定した場所と一致 `Url` [する必要](../reference/manifest/extendedoverrides.md) があります。</span><span class="sxs-lookup"><span data-stu-id="30289-123">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="30289-124">このファイルには、キーボード ショートカットと、キーボード ショートカットが呼び出すアクションが記述されます。</span><span class="sxs-lookup"><span data-stu-id="30289-124">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="30289-125">JSON ファイル内には、2 つの配列があります。</span><span class="sxs-lookup"><span data-stu-id="30289-125">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="30289-126">actions 配列には、呼び出すアクションを定義するオブジェクトが含まれます。ショートカット配列には、キーの組み合わせをアクションにマップするオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="30289-126">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="30289-127">次に例を示します：</span><span class="sxs-lookup"><span data-stu-id="30289-127">Here is an example:</span></span>

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
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="30289-128">JSON オブジェクトの詳細については、「アクション オブジェクトを作成 [する」および](#construct-the-action-objects) 「ショートカット オブジェクトを作成 [する」を参照してください](#construct-the-shortcut-objects)。</span><span class="sxs-lookup"><span data-stu-id="30289-128">For more information about the JSON objects, see [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects).</span></span> <span data-ttu-id="30289-129">ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="30289-129">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="30289-130">この記事では、"Ctrl" の代りで "CONTROL" を使用できます。</span><span class="sxs-lookup"><span data-stu-id="30289-130">You can use "CONTROL" in place of "Ctrl" throughout this article.</span></span>

    <span data-ttu-id="30289-131">後の手順では、アクション自体が作成する関数にマップされます。</span><span class="sxs-lookup"><span data-stu-id="30289-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="30289-132">この例では、後で SHOWTASKPANE をメソッドを呼び出す関数にマップし、HIDETASKPANE をメソッドを呼び出す `Office.addin.showAsTaskpane` 関数にマップ `Office.addin.hide` します。</span><span class="sxs-lookup"><span data-stu-id="30289-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="30289-133">アクションの関数へのマッピングを作成する</span><span class="sxs-lookup"><span data-stu-id="30289-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="30289-134">プロジェクトで、HTML ページによって読み込まれた JavaScript ファイルを要素で開 `<FunctionFile>` きます。</span><span class="sxs-lookup"><span data-stu-id="30289-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="30289-135">JavaScript ファイルで[、Office.actions.associate](/javascript/api/office/office.actions#associate) API を使用して、JSON ファイルで指定した各アクションを JavaScript 関数にマップします。</span><span class="sxs-lookup"><span data-stu-id="30289-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="30289-136">次の JavaScript をファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="30289-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="30289-137">コードについて次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="30289-137">Note the following about the code.</span></span>

    - <span data-ttu-id="30289-138">最初のパラメーターは、JSON ファイルからのアクションの 1 つです。</span><span class="sxs-lookup"><span data-stu-id="30289-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="30289-139">2 番目のパラメーターは、JSON ファイル内のアクションにマップされているキーの組み合わせをユーザーが押すと実行される関数です。</span><span class="sxs-lookup"><span data-stu-id="30289-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="30289-140">この例を続行するには、最初 `'SHOWTASKPANE'` のパラメーターとして使用します。</span><span class="sxs-lookup"><span data-stu-id="30289-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="30289-141">関数の本文では[、Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--)メソッドを使用してアドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="30289-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="30289-142">完了したら、コードは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="30289-142">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="30289-143">2 番目の関数呼び出しを追加して、アクション `Office.actions.associate` `HIDETASKPANE` を[Office.addin.hide を](/javascript/api/office/office.addin#hide--)呼び出す関数にマップします。</span><span class="sxs-lookup"><span data-stu-id="30289-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="30289-144">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="30289-144">The following is an example.</span></span>

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

<span data-ttu-id="30289-145">前の手順に従うと **、Ctrl** + Alt + Up キーと Ctrl + Alt + Down キーを押して、作業ウィンドウの表示を切り替 **えます**。</span><span class="sxs-lookup"><span data-stu-id="30289-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Alt+Up** and **Ctrl+Alt+Down**.</span></span> <span data-ttu-id="30289-146">同じ動作は、Excelアドイン[](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)PnP repo の Officeキーボード ショートカット のサンプルにGitHub。</span><span class="sxs-lookup"><span data-stu-id="30289-146">The same behavior is shown in the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample in the Office Add-ins PnP repo in GitHub.</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="30289-147">詳細と制限</span><span class="sxs-lookup"><span data-stu-id="30289-147">Details and restrictions</span></span>

### <a name="construct-the-action-objects"></a><span data-ttu-id="30289-148">アクション オブジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="30289-148">Construct the action objects</span></span>

<span data-ttu-id="30289-149">次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`actions` します。</span><span class="sxs-lookup"><span data-stu-id="30289-149">Use the following guidelines when specifying the objects in the `actions` array of the shortcuts.json.</span></span>

- <span data-ttu-id="30289-150">プロパティ名 `id` と `name` 必須です。</span><span class="sxs-lookup"><span data-stu-id="30289-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="30289-151">この `id` プロパティは、キーボード ショートカットを使用して呼び出すアクションを一意に識別するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="30289-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="30289-152">プロパティ `name` は、アクションを記述するユーザーフレンドリーな文字列である必要があります。</span><span class="sxs-lookup"><span data-stu-id="30289-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="30289-153">文字 A - Z、a - z、0 ~ 9、および句読点 "-"、"_"、および "+" の組み合わせである必要があります。</span><span class="sxs-lookup"><span data-stu-id="30289-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="30289-154">プロパティは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="30289-154">The `type` property is optional.</span></span> <span data-ttu-id="30289-155">現在は `ExecuteFunction` 型のみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="30289-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="30289-156">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="30289-156">The following is an example.</span></span>

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

<span data-ttu-id="30289-157">ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="30289-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="construct-the-shortcut-objects"></a><span data-ttu-id="30289-158">ショートカット オブジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="30289-158">Construct the shortcut objects</span></span>

<span data-ttu-id="30289-159">次のガイドラインを使用して、オブジェクトの配列内のオブジェクトを指定shortcuts.js`shortcuts` します。</span><span class="sxs-lookup"><span data-stu-id="30289-159">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json.</span></span>

- <span data-ttu-id="30289-160">プロパティ名 `action` 、 `key` および `default` 必須です。</span><span class="sxs-lookup"><span data-stu-id="30289-160">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="30289-161">プロパティの値は `action` 文字列であり、action オブジェクトのプロパティの 1 `id` つと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="30289-161">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="30289-162">プロパティ `default` には、文字 A ~ Z、-z、0 ~ 9、句読点 "-"、"_"、"+" の任意の組み合わせを指定できます。</span><span class="sxs-lookup"><span data-stu-id="30289-162">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="30289-163">(慣例では、これらのプロパティでは小文字は使用されません)。</span><span class="sxs-lookup"><span data-stu-id="30289-163">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="30289-164">プロパティ `default` には、少なくとも 1 つの修飾子キー (Alt、Ctrl、Shift) の名前と、他の 1 つのキーのみを含む必要があります。</span><span class="sxs-lookup"><span data-stu-id="30289-164">The `default` property must contain the name of at least one modifier key (Alt, Ctrl, Shift) and only one other key.</span></span> 
- <span data-ttu-id="30289-165">Shift を唯一の修飾子キーとして使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="30289-165">Shift cannot be used as the only modifier key.</span></span> <span data-ttu-id="30289-166">Shift キーと Alt キーまたは Ctrl キーを組み合わせます。</span><span class="sxs-lookup"><span data-stu-id="30289-166">Combine Shift with either Alt or Ctrl.</span></span>
- <span data-ttu-id="30289-167">Mac では、Command 修飾子キーもサポートしています。</span><span class="sxs-lookup"><span data-stu-id="30289-167">For Macs, we also support the Command modifier key.</span></span>
- <span data-ttu-id="30289-168">Mac の場合、Alt は Option キーにマップされます。</span><span class="sxs-lookup"><span data-stu-id="30289-168">For Macs, Alt is mapped to the Option key.</span></span> <span data-ttu-id="30289-169">このWindows、Command は Ctrl キーにマップされます。</span><span class="sxs-lookup"><span data-stu-id="30289-169">For Windows, Command is mapped to the Ctrl key.</span></span>
- <span data-ttu-id="30289-170">標準キーボードで 2 つの文字が同じ物理キーにリンクされている場合は、プロパティ内の類義語になります。たとえば、Alt+a と Alt+A は同じショートカットなので `default` 、"-" と "_" は同じ物理キーなので、Ctrl + + と Ctrl+ も同じです。 \_</span><span class="sxs-lookup"><span data-stu-id="30289-170">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, Alt+a and Alt+A are the same shortcut, so are Ctrl+- and Ctrl+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="30289-171">"+" 文字は、そのいずれかの側のキーが同時に押された状態を示します。</span><span class="sxs-lookup"><span data-stu-id="30289-171">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="30289-172">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="30289-172">The following is an example.</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

<span data-ttu-id="30289-173">ショートカット JSON の完全なスキーマは、extended-manifest.schema.js[ です](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="30289-173">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="30289-174">キーヒント (Excel ショートカットなどのシーケンシャル キー ショートカットとも呼ばれる) は、Office アドインではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="30289-174">KeyTips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a><span data-ttu-id="30289-175">他のアドインで使用されるキーの組み合わせを回避する</span><span class="sxs-lookup"><span data-stu-id="30289-175">Avoid key combinations in use by other add-ins</span></span>

<span data-ttu-id="30289-176">ユーザーが既に使用しているキーボード ショートカットは多数Office。</span><span class="sxs-lookup"><span data-stu-id="30289-176">There are many keyboard shortcuts that are already in use by Office.</span></span> <span data-ttu-id="30289-177">既に使用されているアドインのキーボード ショートカットを登録しないようにしますが、既存のキーボード ショートカットを上書きしたり、同じキーボード ショートカットを登録した複数のアドイン間の競合を処理する必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="30289-177">Avoid registering keyboard shortcuts for your add-in that are already in use, however there may be some instances where it is necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.</span></span>

<span data-ttu-id="30289-178">競合が発生した場合、ユーザーが最初に競合するキーボード ショートカットを使用しようとすると、ダイアログ ボックスが表示されます。このダイアログに表示されるアクション名は、ファイル内のアクション オブジェクトのプロパティです。 `name` `shortcuts.json`</span><span class="sxs-lookup"><span data-stu-id="30289-178">In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut, note that the action name that is displayed in this dialog is the `name` property in the action object in `shortcuts.json` file.</span></span>

![1 つのショートカットに対して 2 つの異なるアクションを持つ競合モーダルを示す図。](../images/add-in-shortcut-conflict-modal.png)

<span data-ttu-id="30289-180">ユーザーは、キーボード ショートカットで実行する操作を選択できます。</span><span class="sxs-lookup"><span data-stu-id="30289-180">The user can select which action the keyboard shortcut will take.</span></span> <span data-ttu-id="30289-181">選択を行った後、同じショートカットの今後の使用のために基本設定が保存されます。</span><span class="sxs-lookup"><span data-stu-id="30289-181">After making the selection, the preference is saved for future uses of the same shortcut.</span></span> <span data-ttu-id="30289-182">ショートカットの基本設定は、プラットフォームごとにユーザーごとに保存されます。</span><span class="sxs-lookup"><span data-stu-id="30289-182">The shortcut preferences are saved per user, per platform.</span></span> <span data-ttu-id="30289-183">ユーザーが自分の設定を変更する場合は、[教えて]検索ボックスから [Office アドインのショートカット設定のリセット] コマンド **を** 呼び出します。</span><span class="sxs-lookup"><span data-stu-id="30289-183">If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box.</span></span> <span data-ttu-id="30289-184">コマンドを呼び出すと、ユーザーのアドインのショートカット設定がすべてクリアされ、次に競合するショートカットを使用しようとすると、ユーザーに競合ダイアログ ボックスが表示されます。</span><span class="sxs-lookup"><span data-stu-id="30289-184">Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut:</span></span>

![[アドインのショートカットの設定] Excel設定のリセットOfficeを表示するダイアログ ボックスを表示します。](../images/add-in-reset-shortcuts-action.png)

<span data-ttu-id="30289-186">最適なユーザー エクスペリエンスを得る場合は、次の優れた方法を使用して、Excelを最小限にすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="30289-186">For the best user experience, we recommend that you minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="30289-187">キーボード ショートカットのみを使用して、次のパターンを使用します。 \**Ctrl + Shift + Alt +* x\*\*\*、x は他のキーです。 </span><span class="sxs-lookup"><span data-stu-id="30289-187">Use only keyboard shortcuts with the following pattern: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="30289-188">さらにキーボード ショートカットが必要な場合は、[](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)キーボード ショートカットExcel一覧を確認し、アドインで使用しないようにします。</span><span class="sxs-lookup"><span data-stu-id="30289-188">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>
- <span data-ttu-id="30289-189">キーボード フォーカスがアドイン UI 内にある場合 **、Ctrl + Spacebar** と **Ctrl + Shift + F10** は基本的なアクセシビリティ ショートカットとして機能しません。</span><span class="sxs-lookup"><span data-stu-id="30289-189">When the keyboard focus is inside the add-in UI, **Ctrl+Spacebar** and **Ctrl+Shift+F10** will not work as these are essential accessibility shortcuts.</span></span>
- <span data-ttu-id="30289-190">Windows または Mac コンピューターで、検索メニューで [Office アドインのショートカット設定をリセットする] コマンドが使用できない場合は、コンテキスト メニューからリボンをカスタマイズしてリボンにコマンドを手動で追加できます。</span><span class="sxs-lookup"><span data-stu-id="30289-190">On a Windows or Mac computer, if the "Reset Office Add-ins shortcut preferences" command is not available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.</span></span>

## <a name="customize-the-keyboard-shortcuts-per-platform"></a><span data-ttu-id="30289-191">プラットフォームごとにキーボード ショートカットをカスタマイズする</span><span class="sxs-lookup"><span data-stu-id="30289-191">Customize the keyboard shortcuts per platform</span></span>

<span data-ttu-id="30289-192">ショートカットをプラットフォーム固有にカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="30289-192">It's possible to customize shortcuts to be platform-specific.</span></span> <span data-ttu-id="30289-193">次に、次の各プラットフォームのショートカットをカスタマイズするオブジェクトの例を `shortcuts` 示します。 `windows` `mac` `web`</span><span class="sxs-lookup"><span data-stu-id="30289-193">The following is an example of the `shortcuts` object that customizes the shortcuts for each of the following platforms: `windows`, `mac`, `web`.</span></span> <span data-ttu-id="30289-194">ただし、ショートカットごとにショートカット キー `default` が必要です。</span><span class="sxs-lookup"><span data-stu-id="30289-194">Note that you must still have a `default` shortcut key for each shortcut.</span></span>

<span data-ttu-id="30289-195">次の例では、 `default` キーは、指定されていないプラットフォームのフォールバック キーです。</span><span class="sxs-lookup"><span data-stu-id="30289-195">In the following example, the `default` key is the fallback key for any platform that is not specified.</span></span> <span data-ttu-id="30289-196">指定されていない唯一のプラットフォームはWindows、キーはユーザーにのみ `default` 適用Windows。</span><span class="sxs-lookup"><span data-stu-id="30289-196">The only platform not specified is Windows, so the `default` key will only apply to Windows.</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="30289-197">キーボード ショートカット JSON をローカライズする</span><span class="sxs-lookup"><span data-stu-id="30289-197">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="30289-198">アドインが複数のローカライズをサポートしている場合は、アクション オブジェクトのプロパティをローカライズ `name` する必要があります。</span><span class="sxs-lookup"><span data-stu-id="30289-198">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="30289-199">また、アドインがサポートするローカライズの中にアルファベットや異なる書き込みシステムがある場合、キーボードが異なる場合は、ショートカットのローカライズも必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="30289-199">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="30289-200">キーボード ショートカット JSON をローカライズする方法については、「拡張オーバーライドをローカライズする [」を参照してください](../develop/localization.md#localize-extended-overrides)。</span><span class="sxs-lookup"><span data-stu-id="30289-200">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="30289-201">オーバーライドできないブラウザー のショートカット</span><span class="sxs-lookup"><span data-stu-id="30289-201">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="30289-202">Web でカスタム キーボード ショートカットを使用する場合、ブラウザーで使用される一部のキーボード ショートカットをアドインで上書きすることはできません。このリストは進行中の作業です。</span><span class="sxs-lookup"><span data-stu-id="30289-202">When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser cannot be overridden by add-ins. This list is a work in progress.</span></span> <span data-ttu-id="30289-203">上書きできない他の組み合わせを発見した場合は、このページの下部にあるフィードバック ツールを使用してお知らせください。</span><span class="sxs-lookup"><span data-stu-id="30289-203">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="30289-204">Ctrl + N</span><span class="sxs-lookup"><span data-stu-id="30289-204">Ctrl+N</span></span>
- <span data-ttu-id="30289-205">Ctrl + Shift + N</span><span class="sxs-lookup"><span data-stu-id="30289-205">Ctrl+Shift+N</span></span>
- <span data-ttu-id="30289-206">Ctrl + T</span><span class="sxs-lookup"><span data-stu-id="30289-206">Ctrl+T</span></span>
- <span data-ttu-id="30289-207">Ctrl + Shift + T</span><span class="sxs-lookup"><span data-stu-id="30289-207">Ctrl+Shift+T</span></span>
- <span data-ttu-id="30289-208">Ctrl + W</span><span class="sxs-lookup"><span data-stu-id="30289-208">Ctrl+W</span></span>
- <span data-ttu-id="30289-209">Ctrl + PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="30289-209">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="30289-210">次の手順</span><span class="sxs-lookup"><span data-stu-id="30289-210">Next Steps</span></span>

- <span data-ttu-id="30289-211">キーボード ショートカット[Excelアドインの例](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="30289-211">See the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample add-in.</span></span>
- <span data-ttu-id="30289-212">「マニフェストの拡張オーバーライドを処理する」の拡張オーバーライドの操作 [の概要を取得します](../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="30289-212">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
