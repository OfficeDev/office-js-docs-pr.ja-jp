# <a name="ux-design-patterns-for-office-add-ins"></a>Office アドインの UX 設計パターン

Office アドインのユーザーエクスペリエンスの設計では、Office ユーザーが優れたエクスペリエンスを得られるとともに、既定の Office UI 内にシームレスに合致することで、Office の全体的なエクスペリエンスが拡張するようにします。  

当社の UX パターンはコンポーネントで構成されます。 コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。 ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。

Office UI Fabric では、外観も動作も Office の一部のようなコンポーネントを表示します。 Fabric を活用して、Office とシームレスに統合します。 アドインに既存のコンポーネント言語がある場合、Fabric のためにその言語を削除する必要はありません。 Office と統合する際に、それを保持する機会を探します。 スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルやと動作の採用を行う方法を検討してください。

規定のパターンは、共通の顧客シナリオとユーザー エクスペリエンスについての調査に基づくベスト プラクティスのソリューションです。 このようなパターンにより、アドインの設計と開発を素早く始められるとともに、Microsoft とブランド要素の間のバランスを取るためのガイダンスとしても役立ちます。 Microsoft の Fabric デザイン言語のデザイン要素とパートナー固有のブランドの独自性の間のバランスを取る、すっきりしてモダンなユーザー エクスペリエンスによって、ユーザー定着率とアドイン導入率を高められます。

UX パターン テンプレートを使用して、次の作業を行います。

* よくある顧客のシナリオにソリューションとして適用する。
* 設計のベスト プラクティスとして適用する。
* [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) のコンポーネントとスタイルを組み込む。
* Office の既定の UI に視覚的に溶け込むアドインをビルドする。
* UX を概念化し、視覚化する。


## <a name="getting-started"></a>はじめに

パターンは、アドインで共通する主要な操作やエクスペリエンスによって整理されています。 主なグループは次のとおりです。

* [最初の実行エクスペリエンス  (FRE)](../design/first-run-experience-patterns.md)
* [認証](../design/authentication-patterns.md)
* [ナビゲーション](../design/navigation-patterns.md)
* [ブランド化デザイン](../design/branding-patterns.md)

各グループを確認して、ベスト プラクティスを使用してアドインを設計する方法を理解してください。



>注記: この文書全体で示されている画面の例は、**1366x768** の解像度で設計し、表示されています。




## <a name="see-also"></a>関連項目
* [デザイン ツールキット](design-toolkits.md)
* [Office UI Fabric](https://developer.microsoft.com/fabric)
* [Office アドイン開発のベスト プラクティス](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-best-practices)
* [Fabric React の使用を開始する](https://docs.microsoft.com/office/dev/add-ins/design/using-office-ui-fabric-react)
