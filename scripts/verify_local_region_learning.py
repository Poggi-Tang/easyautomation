"""Real local detectors -> recorder point lookup -> saved UIA locations on a disposable Qt app."""
from __future__ import annotations

import argparse
import ctypes
import json
from pathlib import Path
import subprocess
import sys
import time


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output",type=Path,required=True)
    parser.add_argument("--fixture-python",type=Path,default=Path(sys.executable))
    args=parser.parse_args();output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    import uiautomation as auto
    from easy_uiauto import utils
    from easy_uiauto.ctrl import Controller
    from easy_uiauto.mcp import scanner,knowledge,native_windows,dynamic_regions
    from easy_uiauto.mcp.protocol import location_from_xpath
    from easy_uiauto.knowledge_library import KnowledgeLibrary
    from easy_uiauto.region_picker import selection_metadata,normalized_selection
    from easy_uiauto.library_actions import owned_windows

    fixture=subprocess.Popen([str(args.fixture_python),str(Path(__file__).resolve().parents[1]/"tests/fixtures/qt_knowledge_target.py"),
                              "--events",str(output/"events.json")],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding="utf-8")
    target=None
    walk=scanner._walk_controls
    try:
        line=fixture.stdout.readline();assert line.startswith("fixture_ready hwnd="),line
        target=native_windows.matching_windows("",handle=int(line.split("=",1)[1]))[0]
        window=auto.ControlFromHandle(target["handle"]);Controller.activate_window_handle(target["handle"],window)
        time.sleep(.3)
        group=next(c for c,_d in walk(window,12,400) if c.Name=="主页区域" and c.ControlTypeName=="GroupControl")
        file=next(c for c,_d in walk(window,12,400) if c.Name=="文件" and c.ControlTypeName=="MenuItemControl")
        bounds=scanner._window_rect(window);region=scanner._rect(group)
        layout=dict(app_id="local-demo",page_id="manual-main",window_handle=target["handle"],window_class=target["class_name"],
                    window_process_id=target["process_id"],image_size=dict(width=bounds["width"],height=bounds["height"]),restrict_regions=True,
                    regions=[dict(id="work",name="工作区",normalized_rect=knowledge.normalized_rect(region,bounds))])
        def forbidden(*a,**kw):raise AssertionError("Learning traversed the full UIA tree or called the vision API")
        scanner._walk_controls=forbidden;scanner.segment_interface=forbidden;scanner.analyze_control_semantics=forbidden
        points=[];point_lookup=utils.get_control_info
        def point(x,y,**kw):
            points.append((x,y));return point_lookup(x,y,**kw)
        utils.get_control_info=point
        def scan(**kw):
            return scanner.scan_window(target["title"],"","","","acceptance",root=output/"vault",manual_layout=layout,
                                       semantic_mode="native",progress=lambda s:print(s,flush=True),**kw)
        result=scan();directory=knowledge.app_dir("local-demo",output/"vault")
        controls=knowledge.load_index(directory)["controls"]
        assert any(c["control_type"]=="ButtonControl" and c["name"]=="主页动作" for c in controls),controls
        assert any(c.get("observed_value")=="主页区域内容" for c in controls),controls
        assert not any(c["control_type"] in {"GroupControl","PaneControl","ToolBarControl"} for c in controls)
        assert len(points)==result["visual_targets"]
        assert json.loads((output/"events.json").read_text(encoding="utf-8"))==[]
        before=result["observed_control_ids"]
        second=scan();assert before==second["observed_control_ids"]

        # Capture the real menu entry before input, then learn its expanded content.
        store=KnowledgeLibrary(output/"vault")
        capture=output/"menu-before.png";scanner._screenshot_window(bounds).save(capture)
        metadata=selection_metadata(file)
        opener=store.save_captured_control("local-demo",dict(selection={**metadata,"location":location_from_xpath(utils.get_control_xpath(file)),
            "normalized_rect":normalized_selection(metadata["rect"],bounds),"window_scope":{"kind":"main"}},
            capture=dict(image=str(capture),width=bounds["width"],height=bounds["height"],window_name=target["title"],window_class=target["class_name"])))
        Controller.execute_resolved("click",control=file);time.sleep(.2)
        layout["regions"]=[dict(id="all",name="整个界面",normalized_rect=dict(left=0,top=0,right=1,bottom=1))]
        menu=scan(navigation_context=dict(path=[opener["id"]],before_ids=set(before),popup_handles=[w["handle"] for w in owned_windows(target)]))
        index=store.read("local-demo");new=next(c for c in index["controls"] if c["name"]=="新建工程")
        assert new["navigation_path"]==[opener["id"]]
        graph=dynamic_regions.build_tree(index["controls"],index["regions"])
        assert graph["parents"]["control:"+new["id"]].startswith("region:"),graph
        report=dict(ok=True,first=result,repeat=second,expanded=menu,points=len(points),full_tree_enumeration=False,api_requests=0,
                    menu_parent=graph["parents"]["control:"+new["id"]])
        (output/"acceptance.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
        print(json.dumps(dict(ok=True,saved=len(controls),point_calls=len(points),menu_parent=report["menu_parent"])),flush=True)
    finally:
        if target and fixture.poll() is None:
            ctypes.windll.user32.PostMessageW(target["handle"],0x10,0,0)
            try:fixture.wait(timeout=5)
            except subprocess.TimeoutExpired:fixture.terminate();fixture.wait(timeout=5)


if __name__=="__main__":main()
