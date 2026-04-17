import json

from meal_validation_engine import validate_meal_bundle, validate_meal_files


def run_auto_detect_example() -> None:
    result = validate_meal_bundle(
        [
            r"C:\path\to\한성 4월 식수.xlsx",
            r"C:\path\to\한성 4월 일별.xlsx",
            r"C:\path\to\한성 4월 합계.xlsx",
        ]
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


def run_explicit_role_example() -> None:
    result = validate_meal_files(
        meal_path=r"C:\path\to\한성 4월 식수.xlsx",
        daily_path=r"C:\path\to\한성 4월 일별.xlsx",
        summary_path=r"C:\path\to\한성 4월 합계.xlsx",
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    run_auto_detect_example()
    # run_explicit_role_example()

