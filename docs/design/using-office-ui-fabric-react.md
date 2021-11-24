---
title: Office アドインの Fluent UI React
description: アドインで UI FluentをReactするOfficeについて学習します。
ms.date: 11/19/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb53dfcfca644159a10d3b3c1d7bb6911561e58e
ms.sourcegitcommit: b3ddc1ddf7ee810e6470a1ea3a71efd1748233c9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/24/2021
ms.locfileid: "61153464"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a>アドインFluent UI ReactをOfficeする

Fluent UI Reactは、Office を含む幅広い Microsoft 製品にシームレスに適合するエクスペリエンスを構築するように設計された、公式のオープン ソース JavaScript フロントエンド フレームワークです。 CSS-in-JS を使用して、高度にカスタマイズ可能かつ堅牢で最新のアクセス可能な React ベースのコンポーネントを提供します。

> [!NOTE]
> この記事では、アドインのコンテキストFluent UI Reactの使用Office説明します。ただし、さまざまなアプリや拡張機能でもMicrosoft 365使用されます。 詳細については[、「UI](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) web Fluent UI Reactおよびオープン ソースの repo Fluent[を参照してください](https://github.com/microsoft/fluentui)。

この記事では、カスタム コンポーネントを使用して構築されたアドインを作成し、React UI FluentをReactします。

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

![コマンド ライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット。](../images/yo-office-word-react.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。

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

3. アドイン作業ウィンドウを開く場合は、[ホーム] **タブで** [タスクウィンドウの表示] **ボタンを選択** します。 作業ウィンドウの下部にある既定のテキストと [**実行**] ボタンに注意してください。 このチュートリアルの残りの部分では、UI コンポーネントから UX コンポーネントを使用する React コンポーネントを作成して、このテキストとボタンをFluent React。

    ![[タスクウィンドウの表示] リボン ボタンが強調表示された Word アプリケーションと、作業ウィンドウで [実行] ボタンと直前のテキストが強調表示された状態を示すスクリーンショット。](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a>UI コントロールReactを使用するFluentコンポーネントをReact

この時点で、React を使用して構築された非常に基本的な作業ウィンドウ アドインが作成されました。 次の手順に従って、アドイン プロジェクト内で新しい React コンポーネント (`ButtonPrimaryExample`) を作成します。 コンポーネントは、UI `Label` コントロール `PrimaryButton` のコンポーネントとコンポーネントFluent使用React。

1. Yeoman ジェネレーターで作成したプロジェクト フォルダーを開き、**src\taskpane\components** に移動します。
2. そのフォルダーで、**button.tsx** という名前の新しいファイルを作成します。
3. **button.tsx** で、次のコードを追加して `ButtonPrimaryExample` コンポーネントを定義します。

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from '@fluentui/react/lib/components/Button';
import { Label } from '@fluentui/react/lib/components/Label';

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
- 作成に使用Fluentコンポーネント ( 、 、 ) React UI `PrimaryButton` `IButtonProps` `Label` を参照します `ButtonPrimaryExample` 。
- `export class ButtonPrimaryExample extends React.Component` を使用して、新しい `ButtonPrimaryExample` コンポーネントを宣言します。
- ボタンの `onClick` イベントを処理する `insertText` 関数を宣言します。
- `render` 関数で React コンポーネントの UI を定義します。 HTML マークアップは、UI FluentおよびReactを使用し、イベントが発生すると、関数が `Label` `PrimaryButton` `onClick` `insertText` 実行されます。

## <a name="add-the-react-component-to-your-add-in"></a>React コンポーネントをアドインに追加

`ButtonPrimaryExample` **src\components\App.tsx** を開き、次の手順を実行して、アドインにコンポーネントを追加します。

1. **Button.tsx** の参照 `ButtonPrimaryExample` に次のインポート ステートメントを追加します。

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. 次のインポート ステートメントを削除します。

    ```typescript
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

!["Insert text..." を含む Word アプリケーションを示すスクリーンショット。ボタンと直前のテキストが強調表示されます。](../images/word-task-pane-with-react-component.png)

おめでとうございます、UI を使用して作業ウィンドウ アドインReact作成Fluent完了React!

## <a name="see-also"></a>関連項目

- [Word アドイン GettingStartedFabricReact](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Office アドインの Fabric Core](fabric-core.md)
- [Office アドインの UX 設計パターン](ux-design-pattern-templates.md)
