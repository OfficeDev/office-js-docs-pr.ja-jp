---
title: Office アドインの Fluent UI React
description: アドインで UI FluentをReactするOfficeについて学習します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: a71c1a0de64d99a9e52c4ca2a7a948b9c33eb9ed
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076309"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a><span data-ttu-id="61b92-103">アドインFluent UI ReactをOfficeする</span><span class="sxs-lookup"><span data-stu-id="61b92-103">Use Fluent UI React in Office Add-ins</span></span>

<span data-ttu-id="61b92-104">FluentUI Reactは、Microsoft 製品の幅広い範囲 (Office を含む) にシームレスに適合するエクスペリエンスを構築するように設計された、公式のオープン ソース JavaScript フロントエンド フレームワークOffice。</span><span class="sxs-lookup"><span data-stu-id="61b92-104">Fluent UI React is the official open-source JavaScript front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products, including Office.</span></span> <span data-ttu-id="61b92-105">CSS-in-JS を使用して、高度にカスタマイズ可能かつ堅牢で最新のアクセス可能な React ベースのコンポーネントを提供します。</span><span class="sxs-lookup"><span data-stu-id="61b92-105">It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS.</span></span>

> [!NOTE]
> <span data-ttu-id="61b92-106">この記事では、アドインのコンテキストFluent UI Reactの使用Office説明します。ただし、さまざまなアプリや拡張機能でもMicrosoft 365使用されます。</span><span class="sxs-lookup"><span data-stu-id="61b92-106">This article describes the use of Fluent UI React in the context of Office Add-ins. But it is also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="61b92-107">詳細については[、「UI](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) web Fluent UI Reactおよびオープン ソースの repo Fluent[を参照してください](https://github.com/microsoft/fluentui)。</span><span class="sxs-lookup"><span data-stu-id="61b92-107">For more information, see [Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) and the open source repo [Fluent UI Web](https://github.com/microsoft/fluentui).</span></span>

<span data-ttu-id="61b92-108">この記事では、カスタム コンポーネントを使用して構築されたアドインを作成し、React UI FluentをReactします。</span><span class="sxs-lookup"><span data-stu-id="61b92-108">This article describes how to create an add-in that's built with React and uses Fluent UI React components.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="61b92-109">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="61b92-109">Create an add-in project</span></span>

<span data-ttu-id="61b92-110">Office アドイン用の Yeoman ジェネレーターを使用して、React を使用するアドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="61b92-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="61b92-111">前提条件をインストールする</span><span class="sxs-lookup"><span data-stu-id="61b92-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="61b92-112">プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="61b92-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="61b92-113">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="61b92-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="61b92-114">**Choose a script type: (スクリプトの種類を選択)** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="61b92-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="61b92-115">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="61b92-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="61b92-116">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="61b92-116">**Which Office client application would you like to support?**</span></span> `Word`

![コマンド ライン インターフェイスの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット。](../images/yo-office-word-react.png)

<span data-ttu-id="61b92-118">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="61b92-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="61b92-119">試してみる</span><span class="sxs-lookup"><span data-stu-id="61b92-119">Try it out</span></span>

1. <span data-ttu-id="61b92-120">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="61b92-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="61b92-121">以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。</span><span class="sxs-lookup"><span data-stu-id="61b92-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="61b92-122">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="61b92-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="61b92-123">次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="61b92-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="61b92-124">変更を行うには、管理者としてコマンド プロンプトまたはターミナルを実行する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="61b92-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="61b92-125">Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。</span><span class="sxs-lookup"><span data-stu-id="61b92-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="61b92-126">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="61b92-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="61b92-127">Word でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="61b92-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="61b92-128">ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Word が開きます。</span><span class="sxs-lookup"><span data-stu-id="61b92-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="61b92-129">ブラウザー上の Word でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="61b92-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="61b92-130">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="61b92-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="61b92-131">アドインを使用するには、Word on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="61b92-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="61b92-132">アドイン作業ウィンドウを開く場合は、[ホーム] **タブで** [タスクウィンドウの表示] **ボタンを選択** します。</span><span class="sxs-lookup"><span data-stu-id="61b92-132">To open the add-in task pane, on the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="61b92-133">作業ウィンドウの下部にある既定のテキストと [**実行**] ボタンに注意してください。</span><span class="sxs-lookup"><span data-stu-id="61b92-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="61b92-134">このチュートリアルの残りの部分では、UI コンポーネントから UX コンポーネントを使用する React コンポーネントを作成して、このテキストとボタンをFluent React。</span><span class="sxs-lookup"><span data-stu-id="61b92-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fluent UI React.</span></span>

    ![[タスクウィンドウの表示] リボン ボタンが強調表示された Word アプリケーションと、作業ウィンドウで [実行] ボタンと直前のテキストが強調表示された状態を示すスクリーンショット。](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a><span data-ttu-id="61b92-136">UI コントロールReactを使用するFluentコンポーネントをReact</span><span class="sxs-lookup"><span data-stu-id="61b92-136">Create a React component that uses Fluent UI React</span></span>

<span data-ttu-id="61b92-137">この時点で、React を使用して構築された非常に基本的な作業ウィンドウ アドインが作成されました。</span><span class="sxs-lookup"><span data-stu-id="61b92-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="61b92-138">次の手順に従って、アドイン プロジェクト内で新しい React コンポーネント (`ButtonPrimaryExample`) を作成します。</span><span class="sxs-lookup"><span data-stu-id="61b92-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="61b92-139">コンポーネントは、UI `Label` コントロール `PrimaryButton` のコンポーネントとコンポーネントFluent使用React。</span><span class="sxs-lookup"><span data-stu-id="61b92-139">The component uses the `Label` and `PrimaryButton` components from Fluent UI React.</span></span>

1. <span data-ttu-id="61b92-140">Yeoman ジェネレーターで作成したプロジェクト フォルダーを開き、**src\taskpane\components** に移動します。</span><span class="sxs-lookup"><span data-stu-id="61b92-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="61b92-141">そのフォルダーで、**button.tsx** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="61b92-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="61b92-142">**button.tsx** で、次のコードを追加して `ButtonPrimaryExample` コンポーネントを定義します。</span><span class="sxs-lookup"><span data-stu-id="61b92-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // In the click event, write text to the document.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph('Hello Fluent UI React!', Word.InsertLocation.end);
      await context.sync();
    });
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

<span data-ttu-id="61b92-143">このコードは、次の処理を実行します。</span><span class="sxs-lookup"><span data-stu-id="61b92-143">This code does the following:</span></span>

- <span data-ttu-id="61b92-144">`import * as React from 'react';` を使用して、React ライブラリを参照します。</span><span class="sxs-lookup"><span data-stu-id="61b92-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="61b92-145">作成に使用Fluentコンポーネント ( 、 、 ) React UI `PrimaryButton` `IButtonProps` `Label` を参照します `ButtonPrimaryExample` 。</span><span class="sxs-lookup"><span data-stu-id="61b92-145">References the Fluent UI React components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="61b92-146">`export class ButtonPrimaryExample extends React.Component` を使用して、新しい `ButtonPrimaryExample` コンポーネントを宣言します。</span><span class="sxs-lookup"><span data-stu-id="61b92-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="61b92-147">ボタンの `onClick` イベントを処理する `insertText` 関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="61b92-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="61b92-148">`render` 関数で React コンポーネントの UI を定義します。</span><span class="sxs-lookup"><span data-stu-id="61b92-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="61b92-149">HTML マークアップは、UI FluentおよびReactを使用し、イベントが発生すると、関数が `Label` `PrimaryButton` `onClick` `insertText` 実行されます。</span><span class="sxs-lookup"><span data-stu-id="61b92-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fluent UI React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="61b92-150">React コンポーネントをアドインに追加</span><span class="sxs-lookup"><span data-stu-id="61b92-150">Add the React component to your add-in</span></span>

<span data-ttu-id="61b92-151">**src\components\App.tsx** を開いて次の手順を完了することにより、アドインに `ButtonPrimaryExample` コンポーネントを追加します。</span><span class="sxs-lookup"><span data-stu-id="61b92-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="61b92-152">**Button.tsx** の参照 `ButtonPrimaryExample` に次のインポート ステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="61b92-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="61b92-153">次の 2 つのインポート ステートメントを削除します。</span><span class="sxs-lookup"><span data-stu-id="61b92-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="61b92-154">既定の `render()` 関数を、`ButtonPrimaryExample` を使った以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="61b92-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

    ```typescript
    render() {
      return (
        <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
          <ButtonPrimaryExample />
        </HeroList>
        </div>
      );
    }
    ```

4. <span data-ttu-id="61b92-155">**App.tsx** に加えた変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="61b92-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="61b92-156">結果を表示する</span><span class="sxs-lookup"><span data-stu-id="61b92-156">See the result</span></span>

<span data-ttu-id="61b92-157">Word で、**App.tsx** に変更を保存すると、アドイン作業ウィンドウが自動的に更新されます。</span><span class="sxs-lookup"><span data-stu-id="61b92-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="61b92-158">作業ウィンドウ下部の既定のテキストとボタンに、`ButtonPrimaryExample` コンポーネントによって定義された UI が表示されるようになりました。</span><span class="sxs-lookup"><span data-stu-id="61b92-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="61b92-159">[**テキストの挿入**] ボタンを選択してドキュメントにテキストを挿入します。</span><span class="sxs-lookup"><span data-stu-id="61b92-159">Choose the **Insert text...** button to insert text into the document.</span></span>

!["Insert text..." を含む Word アプリケーションを示すスクリーンショット。ボタンと直前のテキストが強調表示されます。](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="61b92-161">おめでとうございます、UI を使用して作業ウィンドウ アドインReact作成Fluent完了React!</span><span class="sxs-lookup"><span data-stu-id="61b92-161">Congratulations, you've successfully created a task pane add-in using React and Fluent UI React!</span></span>

## <a name="see-also"></a><span data-ttu-id="61b92-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="61b92-162">See also</span></span>

- [<span data-ttu-id="61b92-163">Word アドイン GettingStartedFabricReact</span><span class="sxs-lookup"><span data-stu-id="61b92-163">Word Add-in GettingStartedFabricReact</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="61b92-164">Office アドインの Fabric Core</span><span class="sxs-lookup"><span data-stu-id="61b92-164">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="61b92-165">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="61b92-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
