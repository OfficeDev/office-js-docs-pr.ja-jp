---
title: マニフェストの問題を検証し、トラブルシューティングする
description: 以下の方法を使用して、Office アドイン マニフェストを検証します。
ms.date: 11/02/2018
localization_priority: Priority
ms.openlocfilehash: 921adf6f1f398887d96031790facc1fb1425af2b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388377"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="6ed74-103">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="6ed74-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="6ed74-104">以下の方法を使用して、Office アドイン マニフェストの問題を検証し、トラブルシューティングを行います。</span><span class="sxs-lookup"><span data-stu-id="6ed74-104">Use these methods to validate and troubleshoot issues in your Office Add-ins manifest:</span></span> 

- [<span data-ttu-id="6ed74-105">Office アドイン検証ツールを使用してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="6ed74-105">Validate your manifest with the Office Add-in Validator</span></span>](#validate-your-manifest-with-the-office-add-in-validator)   
- [<span data-ttu-id="6ed74-106">XML スキーマと比較してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="6ed74-106">Validate your manifest against the XML schema</span></span>](#validate-your-manifest-against-the-xml-schema)
- [<span data-ttu-id="6ed74-107">Office アドイン用の Yeoman ジェネレーターでマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="6ed74-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [<span data-ttu-id="6ed74-108">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="6ed74-108">Use runtime logging to debug your add-in</span></span>](#use-runtime-logging-to-debug-your-add-in)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a><span data-ttu-id="6ed74-109">Office アドイン検証ツールを使用してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="6ed74-109">Validate your manifest with the Office Add-in Validator</span></span>

<span data-ttu-id="6ed74-110">Office アドインを記述するマニフェスト ファイルが正確かつ完全であることを確認するために、[Office アドイン検証ツール](https://github.com/OfficeDev/office-addin-validator)を使用してマニフェスト ファイルを検証します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-110">To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).</span></span>

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a><span data-ttu-id="6ed74-111">Office アドイン検証ツールを使用してマニフェストを検証するには</span><span class="sxs-lookup"><span data-stu-id="6ed74-111">To use the Office Add-in Validator to validate your manifest</span></span>

1. <span data-ttu-id="6ed74-112">[Node.js](https://nodejs.org/download/) をインストールします。</span><span class="sxs-lookup"><span data-stu-id="6ed74-112">Install [Node.js](https://nodejs.org/download/).</span></span> 

2. <span data-ttu-id="6ed74-113">管理者としてコマンド プロンプト/ターミナルを開き、次のコマンドを使用して Office アドイン検証ツールとその依存関係をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="6ed74-113">Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:</span></span>

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > <span data-ttu-id="6ed74-114">Yo Office が既にインストールされている場合、最新のバージョンにアップグレードすると、検証ツールが依存関係としてインストールされます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-114">If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.</span></span>

3. <span data-ttu-id="6ed74-p101">マニフェストを検証するには、次のコマンドを実行します。MANIFEST.XML をマニフェスト XML ファイルへのパスに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p101">Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.</span></span>

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="6ed74-117">XML スキーマと比較してマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="6ed74-117">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="6ed74-118">マニフェスト ファイルが、使用している要素のすべての名前空間を含む、正しいスキーマに従っていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-118">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="6ed74-119">他のマニフェストのサンプルから要素をコピーした場合は、**適切な名前空間が含まれている**ことも再確認します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-119">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="6ed74-120">マニフェストは、[XML スキーマ定義 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) ファイルと比較して検証できます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-120">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="6ed74-121">XML スキーマの検証ツールを使用して、この検証を実行できます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-121">You can use an XML schema validation tool to perform this validation.</span></span> 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="6ed74-122">コマンド ライン XML スキーマ検証ツールを使用してマニフェストを検証するには</span><span class="sxs-lookup"><span data-stu-id="6ed74-122">To use a command-line XML schema validation tool to validate your manifest</span></span>

1.  <span data-ttu-id="6ed74-123">[tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。</span><span class="sxs-lookup"><span data-stu-id="6ed74-123">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2.  <span data-ttu-id="6ed74-p103">次のコマンドを実行します。`XSD_FILE` をマニフェスト XSD ファイルへのパスに置き換え、`XML_FILE` をマニフェスト XML ファイルへのパスに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p103">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="6ed74-126">Office アドイン用の Yeoman ジェネレーターでマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="6ed74-126">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="6ed74-127">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用して Office アドインを作成した場合、プロジェクトのルート ディレクトリで次のコマンドを実行すると、マニフェスト ファイルが正しいスキーマに従っていることを確認できます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-127">If you've created your Office Add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office), you can ensure that the manifest file follows the correct schema by running the following command within the root directory of your project:</span></span>

```bash
npm run validate
```

![コマンドラインから Yo Office 検証コントロールが実行され、検証の成功結果が生成されたアニメーション gif](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="6ed74-129">この機能にアクセスするには、アドイン プロジェクトが [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office) バージョン 1.1.17 以降を使用して作成されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="6ed74-129">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="6ed74-130">アドインのデバッグにランタイム ログを使用する</span><span class="sxs-lookup"><span data-stu-id="6ed74-130">Use runtime logging to debug your add-in</span></span> 

<span data-ttu-id="6ed74-131">ランタイム ログを使用して、アドインのマニフェストやいくつかのインストール エラーをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-131">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="6ed74-132">この機能は、リソース ID の不一致のような XSD スキーマ検証では検出されないマニフェストの問題を識別して修正するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-132">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="6ed74-133">ランタイム ログは、アドイン コマンドと Excel カスタム関数を実装するアドインのデバッグに特に有効です。</span><span class="sxs-lookup"><span data-stu-id="6ed74-133">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="6ed74-134">ランタイムのログ機能は現在、Office 2016 デスクトップで利用可能です。</span><span class="sxs-lookup"><span data-stu-id="6ed74-134">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="6ed74-135">ランタイムのログを有効にするには</span><span class="sxs-lookup"><span data-stu-id="6ed74-135">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6ed74-p105">ランタイムのログはパフォーマンスに影響します。アドイン マニフェストに関する問題をデバッグする必要がある場合にのみ有効にしてください。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p105">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="6ed74-138">ランタイムのログを有効にするには、以下を実行します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-138">To turn on runtime logging:</span></span>

1. <span data-ttu-id="6ed74-139">Office 2016 デスクトップのビルド **16.0.7019** 以降を実行していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-139">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="6ed74-140">`HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` の下に `RuntimeLogging` レジストリ キーを追加します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-140">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6ed74-141">`Developer` キー (フォルダー) が `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\` の下にまだない場合、次の手順を完了して作成します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-141">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="6ed74-142">**[WEF]** キー (フォルダー) を右クリックし、**[新規]**、**[キー]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-142">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="6ed74-143">新しいキーに **Developer** という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-143">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="6ed74-p106">キーの既定値にログを書き込むファイルの完全なパスを設定します。例については、[EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p106">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6ed74-146">ログ ファイルが書き込まれるディレクトリが既に存在しており、書き込みアクセス許可がある必要があります。</span><span class="sxs-lookup"><span data-stu-id="6ed74-146">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="6ed74-147">レジストリは次の図のようになります。</span><span class="sxs-lookup"><span data-stu-id="6ed74-147">The following image shows what the registry should look like.</span></span> <span data-ttu-id="6ed74-148">この機能を無効にするには、`RuntimeLogging` キーをレジストリから削除します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-148">To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![RuntimeLogging レジストリ キーを追加したレジストリ エディターのスクリーンショット](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="6ed74-150">マニフェストの問題のトラブルシューティングを行うには</span><span class="sxs-lookup"><span data-stu-id="6ed74-150">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="6ed74-151">ランタイムのログを使用してアドインの読み込みに関する問題のトラブルシューティングを行うには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="6ed74-151">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="6ed74-152">テスト用に[アドインをサイドロード](sideload-office-add-ins-for-testing.md)します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-152">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="6ed74-153">ログ ファイルのメッセージ数を最小限に抑えるため、テストするアドインのみをサイドロードすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6ed74-153">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="6ed74-154">何も起こらず、アドインが表示されない (アドイン ダイアログ ボックスにも表示されない) 場合は、ログ ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="6ed74-154">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="6ed74-p108">ログ ファイルでアドインの ID を検索します。ID はマニフェストで定義します。ログ ファイルでは、この ID には `SolutionId` というラベルが付いています。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p108">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="6ed74-p109">次の例のログ ファイルでは、存在しないリソース ファイルを参照しているコントロールが示されています。この例の問題を修正するには、マニフェストの入力ミスを訂正するか、足りないリソースを追加します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p109">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![見つからないリソース ID を指定するエントリが含まれるログ ファイルのスクリーンショット](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="6ed74-160">ランタイムのログに関する既知の問題</span><span class="sxs-lookup"><span data-stu-id="6ed74-160">Known issues with runtime logging</span></span>

<span data-ttu-id="6ed74-p110">混乱を招くメッセージまたは正しく分類されていないメッセージがログ ファイルに書き込まれることがあります。たとえば次のような場合です。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p110">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="6ed74-163">メッセージ "`Medium Current host not in add-in's host list`" に続く "`Unexpected Parsed manifest targeting different host`" は、誤ってエラーとして分類されています。</span><span class="sxs-lookup"><span data-stu-id="6ed74-163">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="6ed74-164">SolutionId が含まれていないメッセージ "`Unexpected Add-in is missing required manifest fields DisplayName`" は、多くの場合、エラーはデバッグ対象のアドインと関係ありません。</span><span class="sxs-lookup"><span data-stu-id="6ed74-164">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="6ed74-p111">`Monitorable` メッセージは、システムの観点からのエラーと予想されます。場合によっては、スキップされたがマニフェスト失敗の原因にはならなかったスペル ミスのある要素のような、マニフェストの問題を示していることがあります。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p111">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="6ed74-167">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="6ed74-167">Clear the Office cache</span></span>

<span data-ttu-id="6ed74-168">リボン ボタンのアイコンのファイル名やアドイン コマンドのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。</span><span class="sxs-lookup"><span data-stu-id="6ed74-168">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="6ed74-169">Windows の場合:</span><span class="sxs-lookup"><span data-stu-id="6ed74-169">For Windows:</span></span>
<span data-ttu-id="6ed74-170">フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-170">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="6ed74-171">Mac の場合: </span><span class="sxs-lookup"><span data-stu-id="6ed74-171">For Mac:</span></span>
<span data-ttu-id="6ed74-172">フォルダー `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` の内容を削除します。</span><span class="sxs-lookup"><span data-stu-id="6ed74-172">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="6ed74-173">iOS の場合: </span><span class="sxs-lookup"><span data-stu-id="6ed74-173">For iOS:</span></span>
<span data-ttu-id="6ed74-p112">アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。</span><span class="sxs-lookup"><span data-stu-id="6ed74-p112">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="6ed74-176">関連項目</span><span class="sxs-lookup"><span data-stu-id="6ed74-176">See also</span></span>

- [<span data-ttu-id="6ed74-177">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="6ed74-177">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="6ed74-178">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="6ed74-178">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="6ed74-179">Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="6ed74-179">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
