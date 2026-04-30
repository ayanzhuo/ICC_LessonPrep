# 生成与维护说明

## 当前状态

目前 `ICC第一次课课件.html` 是手工维护后的渲染结果。  
本文件夹是从 HTML 拆出来的源内容模块，还没有自动生成器。

## 建议下一步

如果继续优化，建议按这个顺序：

1. 先改 `content/05_classroom_script.md`，确认课堂节奏和每段学生动作。
2. 再改 `content/01_content_master.md`，统一课程目标和逐页内容。
3. 如涉及四个项目方向，改 `content/02_project_cards.md`。
4. 如涉及大字号提问页，改 `content/03_question_slides.md`。
5. 如涉及发明日志和查新报告，改 `content/04_competition_materials.md`。
6. 如需要更多故事、反例、追问，改 `content/06_story_and_cases.md`。
7. 如需要图片、视频或已有方案案例，改 `content/07_media_plan.md`。
8. 如涉及页面增删顺序，改 `structure/slide_manifest.yaml`。
9. 如涉及风格，改 `visual/visual_system.md`。
10. 最后再同步生成 HTML / PPT。

## 同步到 HTML / PPT 时的判断

- 大问题页保持低信息密度，只放一个问题和少量视觉元素。
- 内容页不要连续超过 2 页，之后必须插入提问、案例、任务或展示。
- 视频页不展示长链接，页面上只写观看问题；链接放到素材来源页。
- 每个时间段都要有学生产出，不把页面做成单纯讲稿。

## 后续可以做的自动化

可以增加一个生成脚本：

- 输入：`slide_manifest.yaml` + content markdown files + visual system
- 输出：`exports/ICC第一次课课件.html`
- 可选输出：`exports/ICC第一次课课件_可画导入版.pptx`

## 为什么这样拆

原 HTML 已经超过普通手改规模，继续直接在 HTML 里改会有三个问题：

- 内容和样式混在一起，容易改漏。
- 页面顺序和导航锚点容易不同步。
- 后续导 PPT 时要重新整理内容。

拆分后，内容、结构、视觉可以分别判断和修改。
