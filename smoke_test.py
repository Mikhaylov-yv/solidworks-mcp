"""Smoke test against a live SOLIDWORKS: works only in NEW documents it creates
(saved to the temp work dir) and closes them at the end; open documents are not touched."""

import asyncio
import json

import server as s


def show(title, v):
    v = s._clean(v) if not isinstance(v, (dict, str)) else v
    print(f"--- {title}\n{json.dumps(v, ensure_ascii=False, indent=1)[:1500]}", flush=True)


BOX = """
std_planes(model)["front"].Select2(False, 0)
sk = model.SketchManager
sk.InsertSketch(True)
sk.CreateCenterRectangle(0, 0, 0, mm(50), mm(30), 0)
sk.InsertSketch(True)
f = model.FeatureManager.FeatureExtrusion3(True, False, False, 0, 0, mm(20), 0, False, False, False, False,
                                           0, 0, False, False, False, False, True, True, True, 0, 0, False)
model.ClearSelection2(True)
result = f.Name
"""


async def main():
    for d in (await s.documents("list"))["documents"]:  # leftovers of a previous failed run
        if d["title"].startswith("mcp_smoke_box"):
            await s.documents("close", name=d["title"])
    show("status", await s.status())
    part = (await s.documents("new", kind="part"))["created"]["title"]
    show("new part", part)
    show("box", await s.execute_python(BOX, undo_name="коробка"))
    show("tree", await s.feature_tree())
    dims = await s.dimensions()
    show("dimensions", dims)
    depth = next(d["name"] for d in dims["dimensions"] if d["value"] == 20)
    show("set depth 35", await s.dimensions("set", values={depth: 35}))
    show("undo", await s.execute_python("model.EditUndo2(1); model.EditRebuild3()", undo_name=None))
    show("after undo", await s.dimensions())
    show("equation add", await s.equations("add", equation='"Длина" = 120'))
    show("equation link", await s.equations("add", equation=f'"{depth}" = "Длина" / 4'))
    show("equation set", await s.equations("set", index=0, equation='"Длина" = 200'))
    show("dims after equations", await s.dimensions())
    show("equation delete", await s.equations("delete", index=1))
    show("props set", await s.custom_properties("set", properties={"Обозначение": "АБВГ.000001.001",
                                                                 "Наименование": "Коробка тестовая"},
                                                     configuration="active"))
    show("props get", await s.custom_properties(configuration="active"))
    show("mass", await s.mass_properties())
    feat = next(f["name"] for f in (await s.feature_tree())["features"] if f["type"] == "Extrusion")
    show("select_for_user", await s.select_for_user([{"feature": feat}]))
    show("get_selection", await s.get_selection(detailed=True))
    show("view", await s.view("iso", display="shaded_edges"))
    path = str(s.WORK_DIR / "mcp_smoke_box.SLDPRT")
    show("save_as", await s.documents("save_as", path=path))
    drw = await s.create_drawing(views=["standard", "iso"], paper="A3")
    show("create_drawing", drw)
    show("drawing_info", await s.drawing_info())
    show("annotations", await s.insert_model_annotations(types=["dimensions_not_marked"]))
    show("export pdf", await s.export(str(s.WORK_DIR / "mcp_smoke_box.pdf")))
    img = await s.screenshot()
    (s.WORK_DIR / "smoke_screenshot.png").write_bytes(img.data)
    print("screenshot bytes:", len(img.data))
    show("close drawing", await s.documents("close", name=drw["drawing"]["title"]))
    show("close part", await s.documents("close", name=path))


asyncio.run(main())
