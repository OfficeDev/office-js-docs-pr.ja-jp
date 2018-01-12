
# <a name="labsjs-lab-components"></a>LabsJS lab components

Labs.js は、ラボの組み立てに使用できる 4 種類のコンポーネントを提供します。それぞれのコンポーネントの種類が特定の種類のラボ対話をサポートします。たとえば選択問題、自由応答問題、またはレッスンの HTML iFrame に Web ページを表示するアクティビティなどが含まれます。

## <a name="components"></a>コンポーネント

Office Mix は、次の 4 つのラボのコンポーネントの種類をサポートしています。 


-  **Activity component** ( **IActivityComponent**). Presents the user with an activity that must be completed; for example, read a piece of text, watch a video, or interact with a simulation. For more information, see [Labs.Components.ActivityComponentInstance](../../../reference/office-mix/labs.components.activitycomponentinstance.md).
    
-  **Choice component** ( **IChoiceComponent**). Presents the user with a list of choices from which the user must select. Supports single or multiple responses (or no answer at all). Use this component type for true/false, multiple choice, multiple response, or polls. For more information, see [Labs.Components.ChoiceComponentInstance](../../../reference/office-mix/labs.components.choicecomponentinstance.md).
    
-  **Input component** ( **IInputComponent**). Enables free form user input. Use this component type when you want to get responses to questions or math problems from the user, for example, or for other problem types that require text inputs from the user. For more information, see [Labs.Components.InputComponentInstance](../../../reference/office-mix/labs.components.inputcomponentinstance.md).
    
-  **Dynamic component** ( **IDynamicComponent**). Generates other component types at runtime. Use this component type when you have branching questions, for example, where follow-up component types vary depending on a previous user input. This type also enables creating quiz banks or generating problems at runtime. For more information, see [Labs.Components.DynamicComponentInstance](../../../reference/office-mix/labs.components.dynamiccomponentinstance.md).
    

## <a name="additional-resources"></a>その他のリソース



- [Office Mix アドイン](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Office Mix 用 LabsJS ラボの構成と編集](../../powerpoint/office-mix/configuring-and-editing-labsjs-labs-for-office-mix.md)
    
- [チュートリアル:Office Mix 用の最初のラボを作成する](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
