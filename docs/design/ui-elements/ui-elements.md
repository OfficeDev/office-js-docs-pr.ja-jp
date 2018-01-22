# <a name="office-add-in-ui-elements"></a>Office アドイン UI 要素

Office アドインでは、次の 2 種類の UI 要素を使用できます。 

- アドイン コマンド 
- カスタム HTML ベースのインターフェイス

## <a name="add-in-commands"></a>アドイン コマンド
コマンドは、[アドイン XML マニフェスト](../../../docs/develop/define-add-in-commands.md)で定義され、Office UI にネイティブな UX 拡張機能としてレンダーされます。たとえば、Office リボンにボタンを追加するアドイン コマンドを使用できます。 

![アドイン内のアドイン コマンドとカスタム HTML UI 要素が表示されたイメージ](../../images/layouts_addInCommands_v0.03.png)

現時点では、アドインのコマンドはメールのアドインの場合のみサポートされます。詳細については、「[メールのアドイン コマンド](../../outlook/add-in-commands-for-outlook.md)」を参照してください。 

Excel、PowerPoint、Word には、Office リボンの [挿入] タブに、作業ウィンドウ アドインとコンテンツ アドイン用の定義済みのエントリ ポイントがあります。コンテンツ アドインと作業ウィンドウ アドインのカスタム コマンド機能は、まもなく使用できるようになる予定です。 

![Word リボンの [挿入] タブを表示するイメージ](../../images/Word-insert-tab.png)

## <a name="custom-html-based-ui"></a>カスタム HTML ベースの UI
アドインを使用して、カスタム HTML ベースの UI を Office クライアント内に埋め込むことができます。UI の表示に使用可能なコンテナーは、アドインの種類によって異なります。たとえば、作業ウィンドウ アドインはドキュメントの右側のウィンドウにカスタム HTML ベースの UI を表示し、コンテンツ アドインは Office ドキュメント内にカスタム UI を直接表示します。

作成するアドインの種類に関係なく、共通の構築ブロックを使用してカスタム HTML ベースの UI を作成できます。[Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) をこれらの UI 要素に使用して、ご自分のアドインと Office の外観との統一感をもたせることをお勧めします。もちろん、独自の UI 要素を使用して、自分自身のブランドを表現することもできます。

Office UI Fabric には、次の UI 要素が用意されています。

- 文字体裁
- 色
- アイコンの場合
- アニメーション
- 入力コンポーネント
- レイアウト
- ナビゲーション要素

[Github から Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) をダウンロードすることができます。

アドインでOffice UI Fabric を使用する方法を示すサンプルについては、「[Office アドイン Fabric UI サンプル](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)」を参照してください。

**注:**独自のフォントとアイコンのセットを使用する場合は、Office のものと競合しないようにしてください。たとえば、Office のアイコンと同じか似ているアイコンを、自分のアドインで別のものを表すために使用しないでください。 

### <a name="creating-a-customized-color-palette"></a>カスタマイズ カラー パレットを作成する
独自のカラー パレットを使用する場合は、次の点に留意してください。 
 
- ブランドの価値がうまくユーザーに伝わり、アドインのユーザー エクスペリエンスを高めるような色を使用します。
- 色に意味を持たせ、アドイン内で一貫して使用します。たとえば、1 つの色をアクセント カラーとして選び、アドインの視覚的テーマの一貫性を確保します。
- 対話式の要素と非対話式の要素に同じ色を使用しないでください。ナビゲーションやリンクやボタンなど、ユーザーが操作できるアイテムに色を使用する場合は、静的アイテムに同じ色を使用しないでください。
- テキストに色を付けたり、色付きの背景で白い文字を使用したりする場合は、十分なコントラストを確保し、アクセシビリティ ガイドライン (4.5:1 のコントラスト比) を満たすようにしてください。
- 色覚障碍者に配慮し、操作を色だけで識別することは避けてください。

### <a name="theming"></a>テーマ 
Office の配色パターンを採用するにしても、独自のものを使用するにしても、テーマ API を使用することお勧めします。Office テーマ エクスペリエンスの一部としてアドインを作成すると、Office との統一感が増します。


- メール アドインと作業ウィンドウ アドインの場合は、[Context.officeTheme](http://dev.office.com/reference/add-ins/shared/office.context.officetheme) プロパティを使用して Office アプリケーションのテーマに合わせます。この API は、現在 Office 2016 でのみ使用できます。  
- PowerPoint のコンテンツ アドインの場合は、「[PowerPoint アドインで Office テーマを使用する](../../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)」を参照してください。

<!-- Link to theming API docs and Humberto's seed sample. Add screenshot of themed add-in. -->



