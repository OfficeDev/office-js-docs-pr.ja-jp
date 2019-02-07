---
title: React を使用して Excel のアドインを作成する
description: ''
ms.date: 10/19/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 02fd62dca59136fe85ff9b29a6b44576f1ceb8e9
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742367"
---
# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="f28bc-102">React を使用して Excel のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f28bc-102">Build an Excel add-in using React</span></span>

<span data-ttu-id="f28bc-103">この記事では、React と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-103">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f28bc-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="f28bc-104">Prerequisites</span></span>

- [<span data-ttu-id="f28bc-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="f28bc-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="f28bc-106">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="f28bc-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="f28bc-107">Web アプリを作成する</span><span class="sxs-lookup"><span data-stu-id="f28bc-107">Create the web app</span></span>

1. <span data-ttu-id="f28bc-108">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-108">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="f28bc-109">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="f28bc-110">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="f28bc-110">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="f28bc-111">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="f28bc-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="f28bc-112">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`</span><span class="sxs-lookup"><span data-stu-id="f28bc-112">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman ジェネレーター](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="f28bc-114">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="f28bc-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="f28bc-115">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-115">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="f28bc-116">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="f28bc-116">Update the code</span></span>

1. <span data-ttu-id="f28bc-117">コード エディターでファイル **src/styles.less** を開き、次のスタイルをファイルの末尾に追加してファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-117">In your code editor, open the file **src/styles.less**, add the following styles to the end of the file, and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. <span data-ttu-id="f28bc-118">Office アドイン Yeoman ジェネレーターで作成されるプロジェクト テンプレートには、このクイック スタートに不要な React コンポーネントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f28bc-118">The project template that the Office Add-ins Yeoman generator created includes a React component that is not needed for this quick start.</span></span> <span data-ttu-id="f28bc-119">ファイル **src/components/HeroList.tsx** を削除します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-119">Delete the file **src/components/HeroList.tsx**.</span></span>

3. <span data-ttu-id="f28bc-120">ファイル **src/components/Header.tsx** を開き、すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-120">Open the file **src/components/Header.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. <span data-ttu-id="f28bc-121">**Content.tsx** という名前の新しい React コンポーネントを **src/components** フォルダーに作成し、次のコードを追加してファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-121">Create a new React component named **Content.tsx** in the **src/components** folder, add the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. <span data-ttu-id="f28bc-122">ファイル **src/components/App.tsx** を開き、すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-122">Open the file **src/components/App.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    /* global Office, Excel */

    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="f28bc-123">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="f28bc-123">Update the manifest</span></span>

1. <span data-ttu-id="f28bc-124">**manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-124">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="f28bc-125">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f28bc-125">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="f28bc-126">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="f28bc-126">Replace it with your name.</span></span>

3. <span data-ttu-id="f28bc-127">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f28bc-127">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="f28bc-128">これは、**A task pane add-in for Excel** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="f28bc-128">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="f28bc-129">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-129">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="f28bc-130">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="f28bc-130">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="f28bc-131">試してみる</span><span class="sxs-lookup"><span data-stu-id="f28bc-131">Try it out</span></span>

1. <span data-ttu-id="f28bc-132">アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="f28bc-132">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="f28bc-133">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="f28bc-133">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="f28bc-134">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="f28bc-134">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="f28bc-135">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="f28bc-135">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="f28bc-136">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="f28bc-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="f28bc-138">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="f28bc-139">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="f28bc-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="f28bc-141">次の手順</span><span class="sxs-lookup"><span data-stu-id="f28bc-141">Next steps</span></span>

<span data-ttu-id="f28bc-p105">これで完了です。React を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="f28bc-p105">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f28bc-144">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="f28bc-144">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="f28bc-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="f28bc-145">See also</span></span>

* [<span data-ttu-id="f28bc-146">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="f28bc-146">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="f28bc-147">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="f28bc-147">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="f28bc-148">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="f28bc-148">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="f28bc-149">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="f28bc-149">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
