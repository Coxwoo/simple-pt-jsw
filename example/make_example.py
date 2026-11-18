"""simple-pt 예시 덱 — 가상의 증폭기 "Demo Amp" (가짜 주제·가짜 데이터).

    python example/make_example.py               # example/demo_amp.pptx

그림 (example/figures/) 은 figstyle 저장소의 예시 그림 (examples/figures/, 큰 글자 판은 big/) 을 복사한 것이다.
장 구성: 표지 → 목차 → 파트 (Why·Design·Result) 마다 구분 장과 본문 → 요약 → Appendix 색인 → A1·A2.
부제는 질문과 평서문을 섞고, 빨간 줄은 파트를 맺는 장 (Design Target·Robustness·Comparison) 에만 둔다.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "tools"))
from equations import render  # noqa: E402
from simple_pt import Deck  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")         # Windows 콘솔에서도 한글 경고가 깨지지 않게

# 식은 tools/equations.py 로 — example/equations/noise.png·.svg·.tex (여러 줄은 & 자리에 맞춘다)
NOISE_EQ = render([r"\overline{v_{\mathrm{n}}^{2}} &= \frac{4kT\gamma}{g_{\mathrm{m}}}",
                   r"g_{\mathrm{m}} &= \frac{2I_{\mathrm{D}}}{V_{\mathrm{ov}}}"], HERE / "equations" / "noise")


def fig(name, big=False):
    """예시 그림 — big=True 는 figstyle 의 큰 글자 판 (한 줄에 셋 놓을 때)."""
    return HERE / "figures" / ("big" if big else "") / f"{name}.png"


d = Deck("Demo Amp — simple-pt example", parts=["Why", "Design", "Result"])

d.cover("Trading Power for Noise:\nA Demo Low-Noise Amplifier",
        subtitle="simple-pt 양식 예시 — 가상의 증폭기와 가짜 데이터",
        kicker="Team Seminar · Demo Deck",
        name="Your Name", affiliation="Dept. · Lab", date="YYYY-MM-DD",
        notes="안녕하십니까. 오늘은 가상의 저잡음 증폭기로 이 양식의 장 구성을 보여 드리겠습니다.")

d.outline([("Why", "작은 센서 신호 — 증폭기 잡음이 문제"),
           ("Design", "입력단에 전류를 모아 잡음 감소"),
           ("Result", "잡음·전력 목표와 다른 설계 비교")],
          notes="발표는 세 부분입니다.\n\n왜 잡음을 줄여야 하는지, 어떻게 줄였는지, 그리고 결과입니다.")

d.divider("Why", notes="먼저, 왜 잡음을 줄여야 하는지입니다.")

d.figure_bullets("Motivation", "왜 증폭기 잡음부터 줄이나?", part="Why",
                 image=fig("noise_budget"), label="Noise budget (previous design)",
                 takeaway="증폭기 몫 64 % — 다른 블록을 줄여도 효과 작음",
                 bullets=["센서 출력은 수 µV 수준",
                          "증폭기 입력 환산 잡음이 SNR 결정",
                          "전력은 그대로 두고 잡음만 낮추기 어려움"],
                 notes="왜 증폭기 잡음부터 줄여야 할까요?\n\n"
                       "센서 신호가 작아서, 증폭기 잡음이 전체 잡음의 64 % 를 차지합니다.\n\n"
                       "그래서 다른 블록을 줄여도 효과가 작습니다.")

d.figure_bullets("Design Target", "입력 환산 잡음 10 nV/√Hz 이하, 전력 200 µW 이하", part="Why",
                 image=fig("noise_vs_current"), label="Noise vs. input-stage current (sim)",
                 takeaway="150 µA 에서 두 목표 모두 충족",
                 caption="빗금은 전력 한계 밖 (200 µW 넘음)",
                 bullets=["잡음은 입력단 전류의 제곱근에 반비례",
                          "전력 = 1.0 V × (입력단 전류 + 30 µA)",
                          "**150 µA** — 잡음 8.5 nV/√Hz, 전력 180 µW"],
                 key="입력단 150 µA — 잡음·전력 목표 동시 충족",
                 notes="목표는 잡음 10 nV/√Hz 이하, 전력 200 µW 이하입니다.\n\n"
                       "입력단 전류를 늘리면 잡음은 줄지만 전력이 늘어서, 두 목표를 함께 만족하는 150 µA 를 골랐습니다.")

d.divider("Design", notes="다음은, 잡음을 어떻게 줄였는지입니다.")

d.equation("Key Idea", "잡음을 무엇으로 줄이나?", part="Design",
           eq=NOISE_EQ,
           where="k 볼츠만 상수, T 절대 온도, γ 잡음 계수, V_{ov} overdrive 전압",
           cards=[{"title": "Input stage", "body": ["전류 1.5 배·소자 2 배", "g_{m} 증가로 잡음 감소"]},
                  {"title": "Output stage", "body": ["전류 30 µA 그대로", "구동 능력만 유지"]},
                  {"title": "Net effect", "body": ["전력 128 → 180 µW (+40 %)", "잡음 13 → 8.5 nV/√Hz (−35 %)"]}],
           notes="잡음은 입력 소자의 g_m 에 반비례하고, g_m 은 전류에 비례합니다.\n\n"
                 "그래서 전류를 입력단에만 쓰고 출력단은 그대로 두었습니다.\n\n"
                 "전력은 40 % 늘고 잡음은 35 % 줄었습니다.")

d.panels("Robustness", "온도·공정·부하가 바뀌어도 목표 이내", part="Design",
         panels=[{"label": "Temperature (sim)", "image": fig("gain_vs_temp", big=True),
                  "takeaway": "−20–100 °C 에서 72.7 dB 이상"},
                 {"label": "Process corner (sim)", "image": fig("noise_by_corner", big=True),
                  "takeaway": "가장 나쁜 SS 에서도 9.6 nV/√Hz"},
                 {"label": "Load capacitance (sim)", "image": fig("pm_vs_load", big=True),
                  "takeaway": "20 pF 까지 PM 50° 이상"}],
         key="모든 조건에서 목표 이내 — 최악 SS corner 9.6 nV/√Hz (sim)",
         notes="왼쪽은 온도, 가운데는 공정 corner, 오른쪽은 부하입니다.\n\n"
               "가장 나쁜 SS corner 에서도 잡음이 목표 안입니다.")

d.divider("Result", notes="마지막으로 결과입니다.")

d.panels("Results", "이전 설계보다 얼마나 나아졌나?", part="Result",
         panels=[{"label": "Open-loop gain (sim)", "image": fig("gain_vs_freq"),
                  "takeaway": "UGF 9.4 MHz — 이전 설계의 약 3 배"},
                 {"label": "Input-referred noise (sim)", "image": fig("noise_spectrum"),
                  "takeaway": "1 MHz 에서 8.5 nV/√Hz — 목표 이내"}],
         notes="이전 설계보다 얼마나 나아졌을까요?\n\n"
               "이득은 14 dB 늘고 UGF 는 약 3 배가 됐습니다.\n\n"
               "잡음은 1 MHz 에서 8.5 nV/√Hz 로 목표 안입니다.")

d.table("Comparison", "비슷한 전력에서 가장 낮은 잡음 — 대가는 가장 큰 면적", part="Result",
        header=["", "This demo", "Design A [1]", "Design B [2]", "Design C [3]"],
        rows=[["Noise (nV/√Hz)", "**8.5**", "12", "9.8", "15"],
              ["Power (µW)", "180", "150", "260", "120"],
              ["Gain (dB)", "**76**", "70", "72", "65"],
              ["UGF (MHz)", "9.4", "6.0", "11", "4.0"],
              ["Area (mm²)", "0.08", "0.05", "0.07", "0.04"],
              ["Data", "sim", "sim", "meas.", "meas."]],
        widths=[28, 18, 18, 18, 18], highlight=[(0, 1)],
        refs=["[1] Example Ref. A, Demo Journal, 2021.", "[2] Example Ref. B, Demo Conf., 2022.",
              "[3] Example Ref. C, Demo Conf., 2023."],
        key="잡음 8.5 nV/√Hz 로 가장 낮음 — 면적 0.08 mm² 는 가장 큼",
        notes="다른 설계와 비교한 표입니다. 굵은 칸이 이 설계가 다른 설계보다 나은 곳입니다.\n\n"
              "대가는 입력단이 커서 면적이 가장 크다는 점입니다.")

d.summary("Summary", "전류를 입력단에 모아 전력 한계 안에서 잡음 목표 달성", part="Result",
          columns=[{"title": "Problem", "body": ["증폭기가 잡음 예산의 64 %"]},
                   {"title": "Idea", "body": ["입력단에만 전류 집중", "출력단은 그대로"]},
                   {"title": "Result", "body": ["8.5 nV/√Hz·180 µW (sim)", "모든 corner·온도에서 목표 이내"]}],
          closing="남은 일 — 입력단 면적 축소, 실측",
          notes="정리하면, 전류를 입력단에 모아서 전력 한계 안에서 잡음 목표를 달성했습니다.\n\n감사합니다.")

d.appendix_index([("잡음 예산은 어떻게 나눴나", "A1"),
                  ("출력 settling 은 충분한가", "A2")],
                 notes="질문에 맞는 줄의 장 번호를 누르고 Enter — 그 장으로 감 (발표 보기)")

d.table("Noise Budget", "잡음 몫 = (v_{n,i} / v_{n,total})^{2} — 이전 설계 기준", part="Why",
        header=["Source", "Noise (nV/√Hz)", "Share"],
        rows=[["Amplifier", "13.0", "**64 %**"],
              ["Sensor", "7.5", "21 %"],
              ["ADC", "5.1", "10 %"],
              ["Other", "3.6", "5 %"],
              ["**Total**", "**16.3**", "100 %"]],
        widths=[40, 35, 25], width=1100, highlight=[(0, 2)],
        notes="증폭기 몫 64 % — 이전 설계 기준\n\n"
              "잡음은 제곱으로 더함 — 몫은 제곱 비\n\n"
              "센서·ADC 를 줄여도 전체 감소는 작음")

d.figure_bullets("Settling", "계단 입력 뒤 332 ns 안에 ±1 % 이내", part="Result",
                 image=fig("step_settling"), label="Step response (sim)",
                 takeaway="332 ns 에 ±1 % 이내 — 1 µs 샘플 주기에 충분",
                 caption="0.5 V 계단, 부하 2 pF — 빗금은 ±1 % 범위",
                 bullets=["넘침 약 10 %", "다음 샘플 (1 µs) 전에 충분히 정착"],
                 notes="332 ns 에 ±1 % 안 — 1 µs 샘플 주기보다 충분히 짧음\n\n"
                       "넘침 약 10 %\n\n"
                       "0.5 V 계단, 부하 2 pF (sim)")

d.save(HERE / "demo_amp.pptx")
