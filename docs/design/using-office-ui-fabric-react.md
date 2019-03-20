---
title: Office アドインでの Office UI Fabric React の使用
description: ''
ms.date: 02/28/2019
localization_priority: Priority
ms.openlocfilehash: 11bb9daf99d85f1c4551363e9f04056870631378
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691147"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="53c3d-102">Office アドインでの Office UI Fabric React の使用</span><span class="sxs-lookup"><span data-stu-id="53c3d-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="53c3d-p101">Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。React を使ってアドインをビルドする場合は、ユーザー エクスペリエンスを作成するために Fabric React の使用を検討してください。Fabric は、アドインで使用できるボタンやチェックボックスなど、複数の React ベースの UX コンポーネントを提供しています。</span><span class="sxs-lookup"><span data-stu-id="53c3d-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="53c3d-106">アドインで Fabric React コンポーネントの使用を開始するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="53c3d-107">この記事の手順を実行すると、アドインで Fabric Core が使用可能になります。</span><span class="sxs-lookup"><span data-stu-id="53c3d-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="53c3d-108">手順 1 - Office 用の Yeoman ジェネレーターでプロジェクトを作成</span><span class="sxs-lookup"><span data-stu-id="53c3d-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="53c3d-109">Fabric React を使用するアドインを作成するには、Office 用の Yeoman ジェネレーターの使用をお勧めします。</span><span class="sxs-lookup"><span data-stu-id="53c3d-109">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office.</span></span> <span data-ttu-id="53c3d-110">Office 用の Yeoman ジェネレーターは、Office アドインを開発するために必要なプロジェクトのスキャフォールディングとビルドの管理を提供します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-110">The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office Add-in.</span></span>

<span data-ttu-id="53c3d-111">プロジェクトを作成するには、**Windows PowerShell** (コマンド プロンプトではありません) を使用して、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="53c3d-112">必須コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="53c3d-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="53c3d-113">`yo office` を実行して、アドイン用のプロジェクト ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="53c3d-114">Office クライアント アプリケーションを選択するように促されたら、**Word** を選択します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="53c3d-p103">プロジェクト ファイルと同じディレクトリにいることを確認し、`npm start` を実行します。スピナーを表示するブラウザー ウィンドウが自動的に開きます。</span><span class="sxs-lookup"><span data-stu-id="53c3d-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="53c3d-117">[マニフェストをサイドロード](../testing/test-debug-office-add-ins.md)し、アドインのすべての UI を表示します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-117">[Sideload your manifest](../testing/test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="53c3d-118">手順 2 - Fabric React コンポーネントを追加</span><span class="sxs-lookup"><span data-stu-id="53c3d-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="53c3d-p104">次に、アドインに Fabric React コンポーネントを追加します。`ButtonPrimaryExample` と呼ばれる、新しい React コンポーネントを作成します。コンポーネントは Fabric React からの Label と PrimaryButton で構成されています。`ButtonPrimaryExample` を作成するには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="53c3d-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="53c3d-122">Yeoman ジェネレーターで作成したプロジェクト フォルダーを開き、**src\components** に移動します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="53c3d-123">**button.tsx** を作成します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="53c3d-124">**button.tsx** で、次のコードを入力して `ButtonPrimaryExample` コンポーネントを作成します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor() {
    super();
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

<span data-ttu-id="53c3d-125">このコードは、次の処理を実行します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-125">This code does the following:</span></span>

- <span data-ttu-id="53c3d-126">`import * as React from 'react';` を使用して、React ライブラリを参照します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="53c3d-127">`ButtonPrimaryExample` の作成に使用する Fabric コンポーネント (PrimaryButton、IButtonProps、Label) を参照します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="53c3d-128">`export class ButtonPrimaryExample extends React.Component` を使用して、新しいパブリック `ButtonPrimaryExample` コンポーネントを宣言して作成します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="53c3d-129">`onClick` イベントを処理する `insertText` 関数を宣言します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="53c3d-p105">`render` 関数で React コンポーネントの UI を定義します。レンダリングで、コンポーネントの構造を定義します。`render` で、`this.insertText` を使って `onClick` イベントの関連付けを行います。</span><span class="sxs-lookup"><span data-stu-id="53c3d-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="53c3d-133">手順 3 - React コンポーネントをアドインに追加</span><span class="sxs-lookup"><span data-stu-id="53c3d-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="53c3d-134">**src\components\app.tsx** を開いて次の手順を実行することにより、アドインに `ButtonPrimaryExample` を追加します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="53c3d-135">次のインポート ステートメントを追加して、手順 2 で作成した **button.tsx** (ファイル拡張子は必要ありません) から `ButtonPrimaryExample` を参照します。</span><span class="sxs-lookup"><span data-stu-id="53c3d-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="53c3d-136">既定の `render()` 関数を、`<ButtonPrimaryExample />` を使った以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="53c3d-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

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

<span data-ttu-id="53c3d-p106">変更を保存します。アドインを含む開いているすべてのブラウザー インスタンスは、自動的に更新され、`ButtonPrimaryExample` React コンポーネントが表示されます。既定のテキストとボタンが、`ButtonPrimaryExample` で定義されたテキストとプライマリ ボタンに置き換えられることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="53c3d-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>



## <a name="see-also"></a><span data-ttu-id="53c3d-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="53c3d-140">See also</span></span>

- [<span data-ttu-id="53c3d-141">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="53c3d-141">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="53c3d-142">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="53c3d-142">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="53c3d-143">Fabric React のコード サンプルの使用にあたって</span><span class="sxs-lookup"><span data-stu-id="53c3d-143">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="53c3d-144">Office アドイン Fabric UI サンプル (Fabric 1.0 を使用)</span><span class="sxs-lookup"><span data-stu-id="53c3d-144">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="53c3d-145">Office 用の Yeoman ジェネレーター</span><span class="sxs-lookup"><span data-stu-id="53c3d-145">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
