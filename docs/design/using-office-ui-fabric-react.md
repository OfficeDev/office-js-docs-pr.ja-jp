---
title: Office アドインでの Office UI Fabric React の使用
description: Office アドインで Office UI Fabric React を使用する方法について説明します。
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: c738521b82d0cb8f234fd28dc8bb24740962b817
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302601"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="64ba5-103">Office アドインでの Office UI Fabric React の使用</span><span class="sxs-lookup"><span data-stu-id="64ba5-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="64ba5-p101">Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。React を使ってアドインをビルドする場合は、ユーザー エクスペリエンスを作成するために Fabric React の使用を検討してください。Fabric は、アドインで使用できるボタンやチェックボックスなど、複数の React ベースの UX コンポーネントを提供しています。</span><span class="sxs-lookup"><span data-stu-id="64ba5-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="64ba5-107">この記事では、React で構築され Fabric React コンポーネントを使用するアドインを作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span> 

> [!NOTE]
> <span data-ttu-id="64ba5-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) は Fabric React に含まれています。つまり、この記事の手順を完了すると、アドインで Fabric Core にアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="64ba5-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="64ba5-109">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="64ba5-109">Create an Outlook add-in project</span></span>

<span data-ttu-id="64ba5-110">Office アドイン用の Yeoman ジェネレーターを使用して、React を使用するアドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="64ba5-111">前提条件をインストールする</span><span class="sxs-lookup"><span data-stu-id="64ba5-111">Install the prerequisites.</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="64ba5-112">プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="64ba5-112">Create the project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="64ba5-113">Yeoman ジェネレーターを使用して、Word アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-113">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="64ba5-114">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-114">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="64ba5-115">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="64ba5-115">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="64ba5-116">**Choose a script type: (スクリプトの種類を選択)** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="64ba5-116">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="64ba5-117">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="64ba5-117">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="64ba5-118">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="64ba5-118">**Which Office client application would you like to support?**</span></span> `Word`

<span data-ttu-id="64ba5-119">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="64ba5-119">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

### <a name="try-it-out"></a><span data-ttu-id="64ba5-120">試してみる</span><span class="sxs-lookup"><span data-stu-id="64ba5-120">Try it out</span></span>

1. <span data-ttu-id="64ba5-121">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-121">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. <span data-ttu-id="64ba5-122">以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。</span><span class="sxs-lookup"><span data-stu-id="64ba5-122">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="64ba5-123">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="64ba5-123">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="64ba5-124">次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="64ba5-124">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="64ba5-125">Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。</span><span class="sxs-lookup"><span data-stu-id="64ba5-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="64ba5-126">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-126">When you run this command, the local web server will start.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="64ba5-127">Word でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="64ba5-128">ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Word が開きます。</span><span class="sxs-lookup"><span data-stu-id="64ba5-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="64ba5-129">ブラウザー上の Word でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="64ba5-130">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="64ba5-130">When you run this command, the local web server will start.</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="64ba5-131">アドインを使用するには、Word on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="64ba5-131">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="64ba5-132">Word で [**ホーム**] タブを選択し、リボンの [**作業ウィンドウの表示**] ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="64ba5-132">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="64ba5-133">作業ウィンドウの下部にある既定のテキストと [**実行**] ボタンに注意してください。</span><span class="sxs-lookup"><span data-stu-id="64ba5-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="64ba5-134">このチュートリアルの残りの部分では、Fabric React の UX コンポーネントを使用する React コンポーネントを作成して、このテキストとボタンを再定義します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![[作業ウィンドウの表示] リボンのボタンが強調表示され、[実行] ボタンおよびその前のテキストが作業ウィンドウで強調表示された Word アプリケーションのスクリーンショット](../images/word-task-pane-yo-default.png)


## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="64ba5-136">Fabric React を使用する React コンポーネントの作成</span><span class="sxs-lookup"><span data-stu-id="64ba5-136">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="64ba5-137">この時点で、React を使用して構築された非常に基本的な作業ウィンドウ アドインが作成されました。</span><span class="sxs-lookup"><span data-stu-id="64ba5-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="64ba5-138">次の手順に従って、アドイン プロジェクト内で新しい React コンポーネント (`ButtonPrimaryExample`) を作成します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="64ba5-139">このコンポーネントは、 Fabric React の `Label` と `PrimaryButton` コンポーネントを使用します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-139">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="64ba5-140">Yeoman ジェネレーターで作成したプロジェクト フォルダーを開き、**src\taskpane\components** に移動します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-140">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="64ba5-141">そのフォルダーで、**button.tsx** という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="64ba5-142">**button.tsx** で、次のコードを追加して `ButtonPrimaryExample` コンポーネントを定義します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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
      body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
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

<span data-ttu-id="64ba5-143">このコードは、次の処理を実行します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-143">This code does the following:</span></span>

- <span data-ttu-id="64ba5-144">`import * as React from 'react';` を使用して、React ライブラリを参照します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="64ba5-145">`ButtonPrimaryExample` の作成に使用される Fabric コンポーネント (`PrimaryButton`、`IButtonProps`、`Label`) を参照します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-145">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create .</span></span>
- <span data-ttu-id="64ba5-146">`export class ButtonPrimaryExample extends React.Component` を使用して、新しい `ButtonPrimaryExample` コンポーネントを宣言します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-146">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="64ba5-147">ボタンの `onClick` イベントを処理する `insertText` 関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="64ba5-148">`render` 関数で React コンポーネントの UI を定義します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="64ba5-149">HTML マークアップは、Fabric React `Label` と `PrimaryButton` コンポーネントを使用し、`onClick` イベントが発生したときに `insertText` 関数が実行されるように指定します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="64ba5-150">React コンポーネントをアドインに追加</span><span class="sxs-lookup"><span data-stu-id="64ba5-150">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="64ba5-151">**src\components\App.tsx** を開いて次の手順を完了することにより、アドインに `ButtonPrimaryExample` コンポーネントを追加します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="64ba5-152">**Button.tsx** の参照 `ButtonPrimaryExample` に次のインポート ステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="64ba5-153">次の 2 つのインポート ステートメントを削除します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="64ba5-154">既定の `render()` 関数を、`ButtonPrimaryExample` を使った以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="64ba5-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

  4. <span data-ttu-id="64ba5-155">**App.tsx** に加えた変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="64ba5-156">結果を表示する</span><span class="sxs-lookup"><span data-stu-id="64ba5-156">See the result</span></span>

<span data-ttu-id="64ba5-157">Word で、**App.tsx** に変更を保存すると、アドイン作業ウィンドウが自動的に更新されます。</span><span class="sxs-lookup"><span data-stu-id="64ba5-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="64ba5-158">作業ウィンドウ下部の既定のテキストとボタンに、`ButtonPrimaryExample` コンポーネントによって定義された UI が表示されるようになりました。</span><span class="sxs-lookup"><span data-stu-id="64ba5-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="64ba5-159">[**テキストの挿入**] ボタンを選択してドキュメントにテキストを挿入します。</span><span class="sxs-lookup"><span data-stu-id="64ba5-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![[テキストの挿入] ボタンとその前のテキストが強調表示された Word アプリケーションのスクリーンショット](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="64ba5-161">おめでとうございます! これで React および Office UI Fabric React を使用して作業ウィンドウ アドインを作成できました。</span><span class="sxs-lookup"><span data-stu-id="64ba5-161">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span> 

## <a name="see-also"></a><span data-ttu-id="64ba5-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="64ba5-162">See also</span></span>

- [<span data-ttu-id="64ba5-163">Office アドインでの Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="64ba5-163">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="64ba5-164">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="64ba5-164">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="64ba5-165">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="64ba5-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="64ba5-166">Fabric React のコード サンプルの使用にあたって</span><span class="sxs-lookup"><span data-stu-id="64ba5-166">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
