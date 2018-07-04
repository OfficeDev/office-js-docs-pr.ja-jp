# <a name="navigation-patterns"></a><span data-ttu-id="addaf-101">ナビゲーション パターン</span><span class="sxs-lookup"><span data-stu-id="addaf-101">Navigation patterns</span></span>

<span data-ttu-id="addaf-102">アドインの主な機能は、特定のコマンド タイプと限られた画面領域を介してアクセスします。</span><span class="sxs-lookup"><span data-stu-id="addaf-102">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="addaf-103">ナビゲーションは直観的であり、コンテキストを提供し、ユーザーがアドイン全体を簡単に移動できるようにすることが大切です。</span><span class="sxs-lookup"><span data-stu-id="addaf-103">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="addaf-104">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="addaf-104">Best practices</span></span>

| <span data-ttu-id="addaf-105">するべきこと</span><span class="sxs-lookup"><span data-stu-id="addaf-105">Do</span></span>    | <span data-ttu-id="addaf-106">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="addaf-106">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="addaf-107">ユーザーに明確なナビゲーション オプションが表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="addaf-107">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="addaf-108">非標準 UI を使用してナビゲーション プロセスを複雑にすることは避けましょう。</span><span class="sxs-lookup"><span data-stu-id="addaf-108">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="addaf-109">ユーザーがアドインをナビゲートできるように、適宜に次のコンポーネントを使用します。</span><span class="sxs-lookup"><span data-stu-id="addaf-109">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="addaf-110">ユーザーがアドイン内の現在の場所やコンテキストを理解することを難しくするのは避けましょう。</span><span class="sxs-lookup"><span data-stu-id="addaf-110">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="addaf-111">コマンド バー</span><span class="sxs-lookup"><span data-stu-id="addaf-111">command bar</span></span>

<span data-ttu-id="addaf-112">CommandBar は、その下にあるウィンドウ、パネル、または親領域の内容を操作するコマンドを格納するサーフェスです。</span><span class="sxs-lookup"><span data-stu-id="addaf-112">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="addaf-113">オプション機能には、ハンバーガー メニューのアクセス ポイント、検索、およびサイド コマンドが含まれます。</span><span class="sxs-lookup"><span data-stu-id="addaf-113">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![コマンド - デスクトップ作業ウィンドウの仕様](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="addaf-115">タブ バー</span><span class="sxs-lookup"><span data-stu-id="addaf-115">Tab bar</span></span>

<span data-ttu-id="addaf-116">テキストとアイコンが縦に並んだボタンを使用してナビゲーションを表示します。</span><span class="sxs-lookup"><span data-stu-id="addaf-116">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="addaf-117">タブ バーを使用すると、短くてわかりやすいタイトルのタブを使用したナビゲーションを表示できます。</span><span class="sxs-lookup"><span data-stu-id="addaf-117">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![タブ バー - デスクトップ作業ウィンドウの仕様](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="addaf-119">[戻る] ボタン</span><span class="sxs-lookup"><span data-stu-id="addaf-119">Back button</span></span>

<span data-ttu-id="addaf-120">[戻る] ボタンを使用すると、ユーザーはドリルダウン ナビゲーション操作から回復できます。</span><span class="sxs-lookup"><span data-stu-id="addaf-120">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="addaf-121">このパターンを使用すれば、ユーザーは順序のある一連の手順に従えるようになります。</span><span class="sxs-lookup"><span data-stu-id="addaf-121">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![[戻る] ボタン - デスクトップ作業ウィンドウの仕様](../images/add-in-back-button.png)
