---
title: Fluent UI ReactアドインOfficeに含む
description: このアドインで Fluent UI ReactをOfficeする方法について学習します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cb7f04c21a52a2e4a3f271abc56aa325dd2b02fd
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330146"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a>アドインで Fluent UI ReactをOfficeする

Fluent UI Reactは、Office を含む幅広い Microsoft 製品にシームレスに適合するエクスペリエンスを構築するように設計された、公式のオープン ソース JavaScript フロントエンド フレームワークです。 CSS-in-JS を使用して高度にカスタマイズ可能React、堅牢で最新のアクセス可能なコンポーネントを提供します。

> [!NOTE]
> この記事では、アドインのコンテキストでの Fluent UI Reactの使用Office説明します。ただし、さまざまなアプリや拡張機能でもMicrosoft 365使用されます。 詳細については[、「Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react)およびオープンソースの repo Fluent [UI Web」を参照してください](https://github.com/microsoft/fluentui)。

この記事では、このコンポーネントを使用して構築されたアドインを作成し、Fluent UI ReactコンポーネントをReactします。

## <a name="create-an-add-in-project"></a>アドイン プロジェクトの作成

Office アドイン用の Yeoman ジェネレーターを使用して、React を使用するアドイン プロジェクトを作成します。

### <a name="install-the-prerequisites"></a>前提条件をインストールする

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a>プロジェクトを作成する

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using React framework`
- **Choose a script type: (スクリプトの種類を選択)** `TypeScript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Word`

![コマンドライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット](../images/yo-office-word-react.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a>試してみる

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. 以下の手順を実行し、ローカル Web サーバーを起動してアドインのサイドロードを行います。

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。 変更を行うには、管理者としてコマンド プロンプトまたはターミナルを実行する必要がある場合もあります。

    > [!TIP]
    > Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。 このコマンドを実行すると、ローカル Web サーバーが起動します。
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Word でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインが読み込まれた Word が開きます。

        ```command&nbsp;line
        npm start
        ```

    - ブラウザー上の Word でアドインをテストするには、プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

        ```command&nbsp;line
        npm run start:web
        ```

        アドインを使用するには、Word on the web で新しいドキュメントを開き、「[Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。

3. アドイン作業ウィンドウを開く場合は、[ホーム] **タブで** [タスクウィンドウの表示] **ボタンを選択** します。 作業ウィンドウの下部にある既定のテキストと [**実行**] ボタンに注意してください。 このチュートリアルの残りの部分では、Fluent UI から UX コンポーネントを使用する React コンポーネントを作成して、このテキストとボタンを再定義React。

    ![[タスクウィンドウの表示] リボン ボタンが強調表示された Word アプリケーションと作業ウィンドウで強調表示された [実行] ボタンと直前のテキストを示すスクリーンショット](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a>Fluent UI Reactを使用するカスタム コンポーネントを作成React

この時点で、React を使用して構築された非常に基本的な作業ウィンドウ アドインが作成されました。 次の手順に従って、アドイン プロジェクト内で新しい React コンポーネント (`ButtonPrimaryExample`) を作成します。 コンポーネントは Fluent `Label` UI のコンポーネントと `PrimaryButton` コンポーネントを使用React。

1. Yeoman ジェネレーターで作成したプロジェクト フォルダーを開き、**src\taskpane\components** に移動します。
2. そのフォルダーで、**button.tsx** という名前の新しいファイルを作成します。
3. **button.tsx** で、次のコードを追加して `ButtonPrimaryExample` コンポーネントを定義します。

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

このコードは、次の処理を実行します。

- `import * as React from 'react';` を使用して、React ライブラリを参照します。
- 作成に使用Reactコンポーネント ( 、 ) `PrimaryButton` `IButtonProps` を参照 `Label` します `ButtonPrimaryExample` 。
- `export class ButtonPrimaryExample extends React.Component` を使用して、新しい `ButtonPrimaryExample` コンポーネントを宣言します。
- ボタンの `onClick` イベントを処理する `insertText` 関数を宣言します。
- `render` 関数で React コンポーネントの UI を定義します。 HTML マークアップでは、Fluent UI コントロールのコンポーネントとコンポーネントReactを使用し、イベントが発生すると関数 `Label` `PrimaryButton` `onClick` が `insertText` 実行されます。

## <a name="add-the-react-component-to-your-add-in"></a>React コンポーネントをアドインに追加

**src\components\App.tsx** を開いて次の手順を完了することにより、アドインに `ButtonPrimaryExample` コンポーネントを追加します。

1. **Button.tsx** の参照 `ButtonPrimaryExample` に次のインポート ステートメントを追加します。

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. 次の 2 つのインポート ステートメントを削除します。

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. 既定の `render()` 関数を、`ButtonPrimaryExample` を使った以下のコードに置き換えます。

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

4. **App.tsx** に加えた変更を保存します。

## <a name="see-the-result"></a>結果を表示する

Word で、**App.tsx** に変更を保存すると、アドイン作業ウィンドウが自動的に更新されます。 作業ウィンドウ下部の既定のテキストとボタンに、`ButtonPrimaryExample` コンポーネントによって定義された UI が表示されるようになりました。 [**テキストの挿入**] ボタンを選択してドキュメントにテキストを挿入します。

!["Insert text..." を含む Word アプリケーションを示すスクリーンショット。ボタンと直前のテキストが強調表示されている](../images/word-task-pane-with-react-component.png)

おめでとうございます、作業ウィンドウ アドインの作成に成功しました。React Fluent UI React!

## <a name="see-also"></a>関連項目

- [Word アドイン GettingStartedFabricReact](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [ファブリック コア (Office アドイン)](fabric-core.md)
- [Office アドインの UX 設計パターン](ux-design-pattern-templates.md)
