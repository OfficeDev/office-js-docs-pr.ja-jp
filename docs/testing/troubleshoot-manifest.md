---
title: マニフェストの問題を検証し、トラブルシューティングする
description: 以下の方法を使用して、Office アドイン マニフェストを検証します。
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: b6d95f6c5658e33c2f52cc46d7bba686bea5cc44
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617059"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="87ce4-103">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="87ce4-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="87ce4-104">アドインのマニフェスト ファイルを検証して、それが正しくて完全であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="87ce4-105">検証を行うと、アドインをサイドロードするときに「アドイン マニフェストが無効です」というエラーが発生している問題も特定することができます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="87ce4-106">この記事では、複数の方法でマニフェスト ファイルを検証し、アドインに関する問題のトラブルシューティングについて説明します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="87ce4-107">Office アドイン用の Yeoman ジェネレーターでマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="87ce4-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="87ce4-108">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用してアドインを作成した場合は、それを使用してプロジェクトのマニフェスト ファイルを検証することもできます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-108">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="87ce4-109">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-109">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![コマンドラインから Yo Office 検証コントロールが実行され、検証の成功結果が生成されたアニメーション gif](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="87ce4-111">この機能にアクセスするには、アドイン プロジェクトが [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office) バージョン 1.1.17 以降を使用して作成されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="87ce4-111">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-toolbox"></a><span data-ttu-id="87ce4-112">office-toolbox でマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="87ce4-112">Validate your manifest with office-toolbox</span></span>

<span data-ttu-id="87ce4-113">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用せずアドインを作成した場合は、[office-toolbox](https://www.npmjs.com/package/office-toolbox) を使用してマニフェストを検証することもできます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-113">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-toolbox](https://www.npmjs.com/package/office-toolbox).</span></span>

1. <span data-ttu-id="87ce4-114">[Node.js](https://nodejs.org/download/) をインストールします。</span><span class="sxs-lookup"><span data-stu-id="87ce4-114">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="87ce4-115">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-115">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="87ce4-116">`MANIFEST_FILE` をマニフェスト ファイルの名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-116">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-toolbox validate -m MANIFEST_FILE
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="87ce4-117">XML スキーマと比較してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="87ce4-117">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="87ce4-118">マニフェストは、[XML スキーマ定義 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) ファイルと比較して検証することができます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-118">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="87ce4-119">マニフェスト ファイルが、使用している要素のすべての名前空間を含む、正しいスキーマに従っていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-119">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="87ce4-120">他のマニフェストのサンプルから要素をコピーした場合は、**適切な名前空間が含まれている**ことも再確認します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-120">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="87ce4-121">XML スキーマの検証ツールを使用して、この検証を実行できます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-121">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="87ce4-122">コマンド ライン XML スキーマ検証ツールを使用してマニフェストを検証するには</span><span class="sxs-lookup"><span data-stu-id="87ce4-122">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="87ce4-123">[tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。</span><span class="sxs-lookup"><span data-stu-id="87ce4-123">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="87ce4-p105">次のコマンドを実行します。`XSD_FILE` をマニフェスト XSD ファイルへのパスに置き換え、`XML_FILE` をマニフェスト XML ファイルへのパスに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p105">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="87ce4-126">アドインのデバッグにランタイム ログを使用する</span><span class="sxs-lookup"><span data-stu-id="87ce4-126">Use runtime logging to debug your add-in</span></span>

<span data-ttu-id="87ce4-127">ランタイム ログを使用して、アドインのマニフェストやいくつかのインストール エラーをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-127">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="87ce4-128">この機能は、リソース ID の不一致のような XSD スキーマ検証では検出されないマニフェストの問題を識別して修正するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-128">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="87ce4-129">ランタイム ログは、アドイン コマンドと Excel カスタム関数を実装するアドインのデバッグに特に有効です。</span><span class="sxs-lookup"><span data-stu-id="87ce4-129">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="87ce4-130">ランタイムのログ機能は現在、Office 2016 デスクトップで利用可能です。</span><span class="sxs-lookup"><span data-stu-id="87ce4-130">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="87ce4-131">ランタイムのログを有効にするには</span><span class="sxs-lookup"><span data-stu-id="87ce4-131">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="87ce4-p107">ランタイムのログはパフォーマンスに影響します。アドイン マニフェストに関する問題をデバッグする必要がある場合にのみ有効にしてください。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p107">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="87ce4-134">ランタイムのログを有効にするには、以下を実行します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-134">To turn on runtime logging:</span></span>

1. <span data-ttu-id="87ce4-135">Office 2016 デスクトップのビルド **16.0.7019** 以降を実行していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-135">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="87ce4-136">`HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` の下に `RuntimeLogging` レジストリ キーを追加します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-136">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="87ce4-137">`Developer` キー (フォルダー) が `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\` の下にまだない場合、次の手順を完了して作成します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-137">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="87ce4-138">**[WEF]** キー (フォルダー) を右クリックし、**[新規]**、**[キー]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-138">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="87ce4-139">新しいキーに **Developer** という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-139">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="87ce4-p108">キーの既定値にログを書き込むファイルの完全なパスを設定します。例については、[EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p108">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="87ce4-142">ログ ファイルが書き込まれるディレクトリが既に存在しており、書き込みアクセス許可がある必要があります。</span><span class="sxs-lookup"><span data-stu-id="87ce4-142">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="87ce4-p109">レジストリは次の図のようになります。 この機能を無効にするには、`RuntimeLogging` キーをレジストリから削除します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p109">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![RuntimeLogging レジストリ キーを追加したレジストリ エディターのスクリーンショット](http://i.imgur.com/Sa9TyI6.png)

### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="87ce4-146">マニフェストの問題のトラブルシューティングを行うには</span><span class="sxs-lookup"><span data-stu-id="87ce4-146">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="87ce4-147">ランタイムのログを使用してアドインの読み込みに関する問題のトラブルシューティングを行うには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="87ce4-147">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="87ce4-148">テスト用に[アドインをサイドロード](sideload-office-add-ins-for-testing.md)します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-148">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="87ce4-149">ログ ファイルのメッセージ数を最小限に抑えるため、テストするアドインのみをサイドロードすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="87ce4-149">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="87ce4-150">何も起こらず、アドインが表示されない (アドイン ダイアログ ボックスにも表示されない) 場合は、ログ ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="87ce4-150">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="87ce4-p110">ログ ファイルでアドインの ID を検索します。ID はマニフェストで定義します。ログ ファイルでは、この ID には `SolutionId` というラベルが付いています。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p110">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="87ce4-p111">次の例のログ ファイルでは、存在しないリソース ファイルを参照しているコントロールが示されています。この例の問題を修正するには、マニフェストの入力ミスを訂正するか、足りないリソースを追加します。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p111">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![見つからないリソース ID を指定するエントリが含まれるログ ファイルのスクリーンショット](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="87ce4-156">ランタイムのログに関する既知の問題</span><span class="sxs-lookup"><span data-stu-id="87ce4-156">Known issues with runtime logging</span></span>

<span data-ttu-id="87ce4-p112">混乱を招くメッセージまたは正しく分類されていないメッセージがログ ファイルに書き込まれることがあります。たとえば次のような場合です。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p112">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="87ce4-159">メッセージ "`Medium Current host not in add-in's host list`" に続く "`Unexpected Parsed manifest targeting different host`" は、誤ってエラーとして分類されています。</span><span class="sxs-lookup"><span data-stu-id="87ce4-159">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="87ce4-160">SolutionId が含まれていないメッセージ "`Unexpected Add-in is missing required manifest fields DisplayName`" は、多くの場合、エラーはデバッグ対象のアドインと関係ありません。</span><span class="sxs-lookup"><span data-stu-id="87ce4-160">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="87ce4-p113">`Monitorable` メッセージは、システムの観点からのエラーと予想されます。場合によっては、スキップされたがマニフェスト失敗の原因にはならなかったスペル ミスのある要素のような、マニフェストの問題を示していることがあります。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p113">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="87ce4-163">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="87ce4-163">Clear the Office cache</span></span>

<span data-ttu-id="87ce4-164">リボン ボタンのアイコンのファイル名やアドイン コマンドのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="87ce4-164">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="87ce4-165">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="87ce4-165">For Windows:</span></span>
<span data-ttu-id="87ce4-166">フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除する</span><span class="sxs-lookup"><span data-stu-id="87ce4-166">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="87ce4-167">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="87ce4-167">For Mac:</span></span>
<span data-ttu-id="87ce4-168">フォルダー `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` の内容を削除する</span><span class="sxs-lookup"><span data-stu-id="87ce4-168">Delete the content of the folder `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="87ce4-169">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="87ce4-169">For iOS:</span></span>
<span data-ttu-id="87ce4-p114">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="87ce4-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="87ce4-172">関連項目</span><span class="sxs-lookup"><span data-stu-id="87ce4-172">See also</span></span>

- [<span data-ttu-id="87ce4-173">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="87ce4-173">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="87ce4-174">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="87ce4-174">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="87ce4-175">Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="87ce4-175">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
