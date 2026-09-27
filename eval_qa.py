"""
eval_qa.py — AI Q&A 성능 테스트 (25문항)

사용법 (VS Code 터미널에서):
    python eval_qa.py 엑셀파일이_있는_폴더

예)  python eval_qa.py data
     python eval_qa.py .        ← 엑셀 파일이 app.py 와 같은 폴더에 있을 때

필요한 파일 (폴더 안):
    01~04 업무자료 엑셀, 05_Q_A_테스트케이스.xlsx
결과:
    같은 폴더에 Q_A_AI_검증결과.xlsx 가 만들어집니다.
예상 비용: 약 2천~4천원 (질문 25개 + 채점 25개)
"""

import glob
import os
import sys

import pandas as pd

import ai_engine as ai

OLD_ENGINE = {"통과": 15, "부분": 5, "실패": 5}  # 기존 규칙 엔진 검증 결과


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    ok, reason = ai.ai_status()
    if not ok:
        print("AI를 사용할 수 없어요:", reason)
        return

    all_xlsx = glob.glob(os.path.join(folder, "*.xlsx"))
    test_files = [p for p in all_xlsx if "테스트" in os.path.basename(p)]
    data_files = [p for p in all_xlsx
                  if not any(k in os.path.basename(p) for k in ["테스트", "검증", "Q_A", "Q&A"])]
    if not test_files or not data_files:
        print(f"'{folder}' 폴더에서 업무자료 엑셀 또는 테스트케이스 파일을 찾지 못했어요.")
        return

    print("업무자료:", ", ".join(os.path.basename(p) for p in data_files))
    files = [(os.path.basename(p), open(p, "rb").read()) for p in data_files]
    normalized = ai.normalize_excel_files(files)
    for key in ["schedule", "project", "contact", "asset"]:
        df = normalized[key]
        print(f"  {key}: {'인식됨 (' + df.attrs.get('source_file', '') + ')' if df is not None else '없음'}")

    raw = pd.read_excel(test_files[0], header=None)
    header_row = next(i for i in range(len(raw)) if "질문" in [str(v).strip() for v in raw.iloc[i].tolist()])
    cases = pd.read_excel(test_files[0], header=header_row).dropna(subset=["질문"])

    reference_date = ai.default_reference_date(normalized["schedule"]).isoformat()
    print(f"기준일: {reference_date} / 테스트 {len(cases)}문항 시작\n")

    rows = []
    for _, case in cases.iterrows():
        question = str(case["질문"]).strip()
        result = ai.answer_question(
            question,
            schedule_df=normalized["schedule"], project_df=normalized["project"],
            contact_df=normalized["contact"], asset_df=normalized["asset"],
            reference_date=reference_date,
        )
        sources_text = "; ".join(
            f"{s['file']} {s['sheet']} {s.get('row') or ''}행".strip() for s in result.get("sources", [])
        )
        expected_source = f"{case.get('근거 파일', '')} {case.get('근거 위치', '')}"
        try:
            verdict = ai.judge_answer(question, str(case.get("기대 답변", "")), expected_source,
                                      result["answer"], sources_text)
        except Exception as exc:
            verdict = {"verdict": "실패", "reason": f"채점 오류: {exc}"}

        number = case.get("번호", "")
        print(f"[{number}] {verdict['verdict']} · {question}")
        rows.append({
            "번호": number,
            "질문": question,
            "판정": verdict["verdict"],
            "AI 답변": result["answer"],
            "AI 근거": sources_text,
            "기대 답변": case.get("기대 답변", ""),
            "기대 근거": expected_source,
            "채점 사유": verdict.get("reason", ""),
            "엔진": result.get("engine", ""),
        })

    result_df = pd.DataFrame(rows)
    total = len(result_df)
    counts = result_df["판정"].value_counts().to_dict()
    passed, partial, failed = counts.get("통과", 0), counts.get("부분", 0), counts.get("실패", 0)
    old_total = sum(OLD_ENGINE.values())

    summary = pd.DataFrame([
        {"지표": "전체 테스트", "기존 규칙 엔진": old_total, "AI 엔진": total},
        {"지표": "통과", "기존 규칙 엔진": OLD_ENGINE["통과"], "AI 엔진": passed},
        {"지표": "부분", "기존 규칙 엔진": OLD_ENGINE["부분"], "AI 엔진": partial},
        {"지표": "실패", "기존 규칙 엔진": OLD_ENGINE["실패"], "AI 엔진": failed},
        {"지표": "정확 통과율", "기존 규칙 엔진": f"{OLD_ENGINE['통과'] / old_total:.0%}",
         "AI 엔진": f"{passed / total:.0%}" if total else "-"},
        {"지표": "사용 가능률(통과+부분)",
         "기존 규칙 엔진": f"{(OLD_ENGINE['통과'] + OLD_ENGINE['부분']) / old_total:.0%}",
         "AI 엔진": f"{(passed + partial) / total:.0%}" if total else "-"},
    ])

    out_path = os.path.join(folder, "Q_A_AI_검증결과.xlsx")
    with pd.ExcelWriter(out_path) as writer:
        summary.to_excel(writer, sheet_name="요약", index=False)
        result_df.to_excel(writer, sheet_name="검증결과", index=False)

    print("\n" + summary.to_string(index=False))
    usage = ai.usage_summary()
    print(f"\n저장 완료: {out_path}")
    print(f"AI 호출 {usage['calls']}회 · 예상 비용 약 {usage['krw']:,}원")


if __name__ == "__main__":
    main()
