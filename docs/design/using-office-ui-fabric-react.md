---
title: Office アドインでの Office UI Fabric React の使用
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 6013275a9a7a4d5d01f37bbbd268a9258cc82f17
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389285"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a>Office アドインでの Office UI Fabric React の使用

Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。React を使ってアドインをビルドする場合は、ユーザー エクスペリエンスを作成するために Fabric React の使用を検討してください。Fabric は、アドインで使用できるボタンやチェックボックスなど、複数の React ベースの UX コンポーネントを提供しています。

アドインで Fabric React コンポーネントの使用を開始するには、次の手順を実行します。

> [!NOTE]
> この記事の手順を実行すると、アドインで Fabric Core が使用可能になります。

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a>手順 1 - Office 用の Yeoman ジェネレーターでプロジェクトを作成

Fabric React を使用するアドインを作成するには、Office 用の Yeoman ジェネレーターの使用をお勧めします。 Office 用の Yeoman ジェネレーターは、Office アドインを開発するために必要なプロジェクトのスキャフォールディングとビルドの管理を提供します。

プロジェクトを作成するには、**Windows PowerShell** (コマンド プロンプトではありません) を使用して、次の手順を実行します。

1. 必須コンポーネントをインストールします。
2. `yo office` を実行して、アドイン用のプロジェクト ファイルを作成します。
3. Office クライアント アプリケーションを選択するように促されたら、**Word** を選択します。
4. プロジェクト ファイルと同じディレクトリにいることを確認し、`npm start` を実行します。スピナーを表示するブラウザー ウィンドウが自動的に開きます。
5. [マニフェストをサイドロード](..\testing\test-debug-office-add-ins.md)し、アドインのすべての UI を表示します。

## <a name="step-2---add-a-fabric-react-component"></a>手順 2 - Fabric React コンポーネントを追加

次に、アドインに Fabric React コンポーネントを追加します。`ButtonPrimaryExample` と呼ばれる、新しい React コンポーネントを作成します。コンポーネントは Fabric React からの Label と PrimaryButton で構成されています。`ButtonPrimaryExample` を作成するには、次のようにします。

1. Yeoman ジェネレーターで作成したプロジェクト フォルダーを開き、**src\components** に移動します。
2. **button.tsx** を作成します。
3. **button.tsx** で、次のコードを入力して `ButtonPrimaryExample` コンポーネントを作成します。

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

このコードは、次の処理を実行します。

- `import * as React from 'react';` を使用して、React ライブラリを参照します。
- `ButtonPrimaryExample` の作成に使用する Fabric コンポーネント (PrimaryButton、IButtonProps、Label) を参照します。
- `export class ButtonPrimaryExample extends React.Component` を使用して、新しいパブリック `ButtonPrimaryExample` コンポーネントを宣言して作成します。
- `onClick` イベントを処理する `insertText` 関数を宣言します。
- `render` 関数で React コンポーネントの UI を定義します。レンダリングで、コンポーネントの構造を定義します。`render` で、`this.insertText` を使って `onClick` イベントの関連付けを行います。

## <a name="step-3---add-the-react-component-to-your-add-in"></a>手順 3 - React コンポーネントをアドインに追加

**src\components\app.tsx** を開いて次の手順を実行することにより、アドインに `ButtonPrimaryExample` を追加します。

- 次のインポート ステートメントを追加して、手順 2 で作成した **button.tsx** (ファイル拡張子は必要ありません) から `ButtonPrimaryExample` を参照します。

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- 既定の `render()` 関数を、`<ButtonPrimaryExample />` を使った以下のコードに置き換えます。

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

変更を保存します。アドインを含む開いているすべてのブラウザー インスタンスは、自動的に更新され、`ButtonPrimaryExample` React コンポーネントが表示されます。既定のテキストとボタンが、`ButtonPrimaryExample` で定義されたテキストとプライマリ ボタンに置き換えられることに注意してください。



## <a name="see-also"></a>関連項目

- [Office UI Fabric React](https://developer.microsoft.com/fabric)
- [Fabric React のコード サンプルの使用にあたって](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [UX 設計パターン (Fabric 2.6.1 を使用)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office アドイン Fabric UI サンプル (Fabric 1.0 を使用)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office 用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)
