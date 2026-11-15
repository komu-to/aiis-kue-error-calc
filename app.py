"""
Главная точка входа приложения «Генератор таблиц погрешностей АИИС КУЭ».
Поддерживает:
1. Графический интерфейс для Windows 10/11 (по умолчанию):
   python app.py

2. Консольный режим (CLI) для пакетной автоматизации и скриптов:
   python app.py --cli --tt 0.5S --th 0.5 --meter "0.2S / 0.5" --dt 20 --out result.xlsx
"""

import sys
import argparse
from core import generate_group_table, PRESET_PROJECT_GROUPS
from exporters import (
    export_to_excel,
    export_to_docx,
    get_tsv_text,
    make_export_filename,
    export_table51_to_excel,
    export_table51_to_docx,
    get_table51_tsv_text,
)


def parse_args():
    parser = argparse.ArgumentParser(description="Расчет погрешностей ИК АИИС КУЭ (ГОСТ Р 8.596 / МИ 3000)")
    parser.add_argument("--cli", action="store_true", help="Запустить в консольном режиме без GUI")
    parser.add_argument("--ik", type=str, default="1", help="Номер или наименование ИК")
    parser.add_argument("--tt", type=str, default="0.5S", help="Класс точности ТТ (0.2S, 0.5S, 0.2, 0.5, 'Без ТТ')")
    parser.add_argument("--th", type=str, default="0.5", help="Класс точности ТН (0.2, 0.5, 'Без ТН')")
    parser.add_argument("--meter", type=str, default="0.2S / 0.5", help="Класс точности счетчика (0.2S / 0.5, 0.5S / 1.0, 1.0 / 1.0, 1.0 / 2.0)")
    parser.add_argument("--dt", type=float, default=20.0, help="Предельное отклонение температуры Δt, °C (по умолчанию 20.0)")
    parser.add_argument("--dL", type=float, default=0.25, help="Потери напряжения в линии связи ТН-счетчик, %% (по умолчанию 0.25)")
    parser.add_argument("--mag", choices=["gost", "all", "all_points", "none", "custom"], default="gost", help="Режим влияния магнитной индукции: gost (по ГОСТ 31819 при Iном, cos=1), all (во всех точках), none, custom")
    parser.add_argument("--mag-val", type=float, default=0.5, help="Значение δми в %% для пользовательского режима (по умолчанию 0.5)")
    parser.add_argument("--all-groups", action="store_true", help="Сгенерировать все типовые группы проекта")
    parser.add_argument("--mode", choices=["summary", "full", "table51"], default="full", help="Формат таблицы: summary (15 граф), full (29 граф) или table51 (матрица Таблицы 5.1 МИ)")
    parser.add_argument("--table51", action="store_true", help="Экспортировать в формат Таблицы 5.1 (матрица по cos phi и токам)")
    parser.add_argument("--val-type", choices=["osn", "rab"], default="osn", help="Тип выводимой погрешности для Таблицы 5.1: osn (δW осн, по умолчанию) или rab (δW раб)")
    parser.add_argument("--out", type=str, help="Путь для сохранения файла (.xlsx или .docx)")
    return parser.parse_args()


def run_cli(args):
    print("=== Расчет погрешностей ИК АИИС КУЭ (Консольный режим) ===")
    mag_mode = "all_points" if args.mag == "all" else args.mag

    if args.all_groups:
        print(f"Генерация всех типовых групп ({len(PRESET_PROJECT_GROUPS)} групп)...")
        all_rows = []
        for g in PRESET_PROJECT_GROUPS:
            rows = generate_group_table(
                ik_num=g["ik"],
                tt_class=g["tt"],
                th_class=g["th"],
                meter_class=g["meter"],
                delta_t=args.dt,
                delta_L=g["delta_L"],
                magnetic_mode=mag_mode,
                custom_dmi_val=args.mag_val
            )
            all_rows.extend(rows)
        target_rows = all_rows
    else:
        print(f"Канал [{args.ik}]: ТТ {args.tt}, ТН {args.th}, Счетчик {args.meter}, Δt={args.dt}°C, δL={args.dL}%, Магнитный режим={mag_mode}")
        target_rows = generate_group_table(
            ik_num=args.ik,
            tt_class=args.tt,
            th_class=args.th,
            meter_class=args.meter,
            delta_t=args.dt,
            delta_L=args.dL,
            magnetic_mode=mag_mode,
            custom_dmi_val=args.mag_val
        )

    print(f"Рассчитано строк: {len(target_rows)}")

    is_table51 = args.table51 or args.mode == "table51"

    if args.out:
        out_path = args.out
        if out_path.lower() in ["auto", "auto.xlsx", "auto.docx", ".xlsx", ".docx", "xlsx", "docx"]:
            ext = "docx" if "docx" in out_path.lower() else "xlsx"
            base_name = make_export_filename(
                ik_str=args.ik,
                tt_str=args.tt,
                th_str=args.th,
                meter_str=args.meter,
                delta_t=args.dt,
                ext=ext,
                is_all_groups=args.all_groups
            )
            out_path = f"Таблица_5.1_δW_{args.val_type}-{base_name}" if is_table51 else base_name

        if is_table51:
            val_sym = "δWосн" if args.val_type == "osn" else "δWраб"
            title = f"Таблица 5.1 — Пределы допускаемой основной погрешности ИК, ± {val_sym} (%)"
            if out_path.lower().endswith(".xlsx"):
                export_table51_to_excel(target_rows, out_path, val_type=args.val_type, title=title)
                print(f"Успешно экспортировано в Excel (Таблица 5.1): {out_path}")
            elif out_path.lower().endswith(".docx"):
                export_table51_to_docx(target_rows, out_path, val_type=args.val_type, title=title)
                print(f"Успешно экспортировано в Word (Таблица 5.1): {out_path}")
            else:
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(get_table51_tsv_text(target_rows, val_type=args.val_type))
                print(f"Успешно сохранен TSV (Таблица 5.1): {out_path}")
        else:
            if out_path.lower().endswith(".xlsx"):
                title = f"Таблица погрешностей ИК {args.ik}" if not args.all_groups else f"Таблица погрешностей всех групп ИК"
                export_to_excel(target_rows, out_path, mode=args.mode, title=title)
                print(f"Успешно экспортировано в Excel: {out_path}")
            elif out_path.lower().endswith(".docx"):
                title = f"Таблица погрешностей ИК {args.ik}" if not args.all_groups else f"Таблица погрешностей всех групп ИК"
                export_to_docx(target_rows, out_path, mode=args.mode, title=title)
                print(f"Успешно экспортировано в Word: {out_path}")
            else:
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(get_tsv_text(target_rows, mode=args.mode))
                print(f"Успешно сохранен TSV: {out_path}")
    else:
        print("\nПервые строки с точкой I=100%:")
        for r in target_rows:
            if r["i1"] in ["100", "100-120"]:
                print(f"  I1={r['i1']:5s} cos={r['cos_phi']:4s} | δми={r['dmi_act']:4s}% | dW_осн_акт={r['dW_osn_act']:4s}% | dW_раб_акт={r['dW_rab_act']:4s}%")


def main():
    args = parse_args()
    if args.cli or args.out or args.all_groups:
        run_cli(args)
    else:
        from gui import main as gui_main
        gui_main()


if __name__ == "__main__":
    main()
