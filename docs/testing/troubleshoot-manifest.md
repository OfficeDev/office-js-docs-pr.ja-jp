---
title: Office アドインのマニフェストを検証する
description: XML スキーマやその他のツールを使用して Office アドインのマニフェストを検証する方法について説明します。
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 9cd1c353d6f73decb5e39df96cf66da5912b8f9c
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163914"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="b072c-103">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="b072c-103">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="b072c-104">アドインのマニフェスト ファイルを検証して、それが正しくて完全であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="b072c-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="b072c-105">検証を行うと、アドインをサイドロードするときに「アドイン マニフェストが無効です」というエラーが発生している問題も特定することができます。</span><span class="sxs-lookup"><span data-stu-id="b072c-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="b072c-106">この記事では、マニフェスト ファイルを検証するための複数の方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="b072c-106">This article describes multiple ways to validate the manifest file.</span></span>

> [!NOTE]
> <span data-ttu-id="b072c-107">ランタイム ログを使用してアドインのマニフェストでの問題をトラブルシューティングする方法の詳細については、「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b072c-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="b072c-108">Office アドイン用の Yeoman ジェネレーターでマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="b072c-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="b072c-109">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用してアドインを作成した場合は、それを使用してプロジェクトのマニフェスト ファイルを検証することもできます。</span><span class="sxs-lookup"><span data-stu-id="b072c-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="b072c-110">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="b072c-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![コマンドラインから Yo Office 検証コントロールが実行され、検証の成功結果が生成されたアニメーション gif](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="b072c-112">この機能にアクセスするには、アドイン プロジェクトが [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office) バージョン 1.1.17 以降を使用して作成されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b072c-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="b072c-113">office-addin-manifest を使用してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="b072c-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="b072c-114">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用せずアドインを作成した場合は、[office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest) を使用してマニフェストを検証することもできます。</span><span class="sxs-lookup"><span data-stu-id="b072c-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="b072c-115">[Node.js](https://nodejs.org/download/) をインストールします。</span><span class="sxs-lookup"><span data-stu-id="b072c-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="b072c-116">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="b072c-116">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="b072c-117">`MANIFEST_FILE` をマニフェスト ファイルの名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b072c-117">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="b072c-118">このコマンドを実行すると、「コマンドの構文が無効です」というエラーメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b072c-118">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="b072c-119">(`validate` コマンドが認識されないため)、次のコマンドを実行してマニフェストを検証します (`MANIFEST_FILE` をマニフェスト ファイル名で置き換えます)。</span><span class="sxs-lookup"><span data-stu-id="b072c-119">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    >
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="b072c-120">XML スキーマと比較してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="b072c-120">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="b072c-121">マニフェストは、[XML スキーマ定義 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) ファイルと比較して検証することができます。</span><span class="sxs-lookup"><span data-stu-id="b072c-121">You can validate the manifest file against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) files.</span></span> <span data-ttu-id="b072c-122">マニフェスト ファイルが、使用している要素のすべての名前空間を含む、正しいスキーマに従っていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="b072c-122">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="b072c-123">他のマニフェストのサンプルから要素をコピーした場合は、**適切な名前空間が含まれている**ことも再確認します。</span><span class="sxs-lookup"><span data-stu-id="b072c-123">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="b072c-124">XML スキーマの検証ツールを使用して、この検証を実行できます。</span><span class="sxs-lookup"><span data-stu-id="b072c-124">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="b072c-125">コマンド ライン XML スキーマ検証ツールを使用してマニフェストを検証するには</span><span class="sxs-lookup"><span data-stu-id="b072c-125">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="b072c-126">[tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。</span><span class="sxs-lookup"><span data-stu-id="b072c-126">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="b072c-p106">次のコマンドを実行します。`XSD_FILE` をマニフェスト XSD ファイルへのパスに置き換え、`XML_FILE` をマニフェスト XML ファイルへのパスに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b072c-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="b072c-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="b072c-129">See also</span></span>

- [<span data-ttu-id="b072c-130">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="b072c-130">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="b072c-131">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="b072c-131">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="b072c-132">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="b072c-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="b072c-133">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="b072c-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="b072c-134">Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="b072c-134">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)