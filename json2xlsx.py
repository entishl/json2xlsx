import json
import pandas as pd
import os
import sys
import openpyxl
# Define the 9 function types we are interested in for aggregation
# This ensures consistent column order and handles cases where a character
# might not have all types.
AGGREGATED_FUNCTION_TYPES = [
    "优越",             # IncElementDmg
    "攻击",             # StatAtk
    "蓄力速度",         # StatChargeTime
    "弹夹",             # StatAmmoLoad
    "蓄力伤害",         # StatChargeDamage
    "防御",             # StatDef
    "暴击率",           # StatCritical
    "命中率",           # StatAccuracyCircle
    "暴击伤害"          # StatCriticalDamage
]


def convert_json_data_to_excel(data, output_excel_path):
    """
    Converts the parsed JSON data (dictionary) into an Excel file
    with multiple sheets, aggregating specific equipment stats.
    """
    # Map old English function types to new Chinese names
    OLD_FUNCTION_TYPE_NAMES = [
        "IncElementDmg", "StatAtk", "StatChargeTime", "StatAmmoLoad",
        "StatChargeDamage", "StatDef", "StatCritical",
        "StatAccuracyCircle", "StatCriticalDamage"
    ]
    FUNCTION_TYPE_MAP = dict(zip(OLD_FUNCTION_TYPE_NAMES,
                                 AGGREGATED_FUNCTION_TYPES))

    try:
        # --- Sheet 1: 基本信息 (Combined with Cube Info) ---
        player_name = data.get("name", "N/A")
        synchro_level = data.get("synchroLevel", "N/A")
        
        target_cubes_levels = {
            "遗迹巨熊魔方": "N/A",
            "战术巨熊魔方": "N/A"
        }
        
        for cube_name, details in data.get("cubes", {}).items():
            if cube_name in target_cubes_levels:
                target_cubes_levels[cube_name] = details.get("cube_level", "N/A")  # noqa: E501
        
        combined_basic_info_data = {
            "玩家": [player_name],
            "同步等级": [synchro_level],
            "遗迹巨熊魔方": [target_cubes_levels["遗迹巨熊魔方"]],
            "战术巨熊魔方": [target_cubes_levels["战术巨熊魔方"]]
        }
        df_basic_info = pd.DataFrame(combined_basic_info_data)

        # --- Sheet 3: 角色信息 ---
        characters_list = []
        # pylint: disable=line-too-long
        for element_type, characters_in_element in data.get("elements", {}).items():  # noqa: E501
            for char_name, char_details in characters_in_element.items():
                char_data = {
                    "元素类型": element_type,
                    "角色": char_name,
                    "Name Code": char_details.get("name_code", "N/A"),
                    "ID": char_details.get("id", "N/A"),
                    "Priority": char_details.get("priority", "N/A"),
                    "技能1等级": char_details.get("skill1_level", "N/A"),
                    "技能2等级": char_details.get("skill2_level", "N/A"),
                    "爆裂技能等级": char_details.get("skill_burst_level", "N/A"),
                    "收藏品稀有度": char_details.get("item_rare", "N/A"),
                    "收藏品等级": char_details.get("item_level", "N/A"),
                    "突破次数": char_details.get("limit_break", "N/A")
                }
                
                # Initialize sums for aggregated stats
                aggregated_stats = {func_type: 0.0 for func_type in AGGREGATED_FUNCTION_TYPES}  # noqa: E501
                
                equipments = char_details.get("equipments", {})
                # Iterate through "0", "1", "2", "3" if they exist
                for slot_key in equipments:
                    slot_data = equipments.get(slot_key, [])
                    for effect in slot_data:
                        func_type_from_json = effect.get("function_type")
                        # Keep as string initially for checking
                        func_value_str = effect.get("function_value")
                        
                        # Get the mapped Chinese name
                        mapped_func_type = FUNCTION_TYPE_MAP.get(
                            func_type_from_json
                        )
                        
                        if (mapped_func_type and
                                mapped_func_type in aggregated_stats):
                            try:
                                # Convert function_value to float for summation
                                func_value_float = float(func_value_str)
                                aggregated_stats[mapped_func_type] += \
                                    func_value_float
                            except (ValueError, TypeError):
                                # pylint: disable=line-too-long
                                part1 = "  Warning: Non-numeric or "
                                part1_cont = "missing function_value "
                                part2_val = f"'{func_value_str}' for "
                                # Use original func_type_from_json for warning
                                # if mapped_func_type is None
                                part2_type_display = (func_type_from_json or
                                                      "Unknown Type")
                                part2_type = (
                                    f"'{part2_type_display}' (mapped to "
                                    f"'{mapped_func_type}') in "
                                )
                                part3 = f"character '{char_name}', element "
                                part4 = f"'{element_type}'. "
                                part4_cont = "Skipping this effect value."
                                warning_msg = (part1 + part1_cont +
                                               part2_val + part2_type +
                                               part3 + part4 + part4_cont)
                                print(warning_msg)
                        # Optional: Add a warning if func_type_from_json
                        # is not in FUNCTION_TYPE_MAP
                        # else:
                        #     if func_type_from_json not in OLD_FUNCTION_TYPE_NAMES and func_type_from_json: # noqa E501
                        #         print(f"  Info: Unmapped function_type '{func_type_from_json}' encountered for character '{char_name}'. Skipping.") # noqa E501
                
                # Add aggregated stats to char_data
                for func_type, total_value in aggregated_stats.items():
                    # Round to a reasonable number of decimal places
                    # if desired, e.g., 2
                    char_data[func_type] = round(total_value, 2)
                    
                characters_list.append(char_data)
        
        df_characters = pd.DataFrame(characters_list) if characters_list else pd.DataFrame([{"元素类型": "无数据"}]) # noqa E501
        
        # If characters_list is not empty, add "玩家" column and
        # ensure all AGGREGATED_FUNCTION_TYPES columns exist
        if characters_list:
            # Insert "玩家" column at the beginning
            df_characters.insert(0, "玩家", player_name)
            
            for func_type_col in AGGREGATED_FUNCTION_TYPES:
                if func_type_col not in df_characters.columns:
                    df_characters[func_type_col] = 0.0  # Add missing columns with 0.0 # noqa: E501
        elif (not df_characters.empty and
              "元素类型" in df_characters.columns and
              df_characters.iloc[0]["元素类型"] == "无数据"):
            # Handle empty characters_list with placeholder row
            df_characters.insert(0, "玩家", player_name)
            # Ensure AGGREGATED_FUNCTION_TYPES columns exist
            for func_type_col in AGGREGATED_FUNCTION_TYPES:
                if func_type_col not in df_characters.columns:
                    df_characters[func_type_col] = 0.0
        
        # --- Write to Excel file ---
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            df_basic_info.to_excel(writer, sheet_name="基本信息", index=False)
            df_characters.to_excel(writer, sheet_name="角色信息", index=False)
        
        print(f"成功转换: {output_excel_path}")
        return True

    except KeyError as e:
        print(
            f"  错误: 文件 {os.path.basename(output_excel_path)} "
            f"的源JSON缺少关键字段 '{e}'。跳过此文件。"
        )  # noqa: E501
        return False
    except Exception as e:
        print(
            f"  错误: 处理文件生成 {os.path.basename(output_excel_path)} "
            f"时发生意外错误: {e}。跳过此文件。"
        )  # noqa: E501
        return False


def main():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Running in a PyInstaller bundle (frozen)
        # sys.executable is the path to the exe
        script_dir = os.path.dirname(sys.executable)
    else:
        # Running as a normal Python script
        script_dir = os.path.dirname(os.path.abspath(__file__))    
    input_folder = os.path.join(script_dir, "input")
    output_folder = os.path.join(script_dir, "output")

    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print(f"已创建输入目录: {input_folder}")
        print("请将JSON文件放入input目录后重新运行脚本。")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"已创建输出目录: {output_folder}")

    json_files_processed = 0
    total_json_files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith(".json")
    ]

    if not total_json_files:
        print(f"\n在 '{input_folder}' 目录中未找到JSON文件。")
        return

    for filename in total_json_files:
        input_json_path = os.path.join(input_folder, filename)
        base_filename = os.path.splitext(filename)[0]
        output_excel_path = os.path.join(output_folder, f"{base_filename}.xlsx")  # noqa: E501
        
        print(f"\n正在处理: {filename} ...")
        
        try:
            with open(input_json_path, 'r', encoding='utf-8') as f:
                data = json.load(f) 
            
            if convert_json_data_to_excel(data, output_excel_path):
                json_files_processed += 1

        except json.JSONDecodeError:
            print(f"  错误: 文件 {filename} 不是有效的JSON格式。跳过。")
        except FileNotFoundError:  # Should not happen if os.listdir works
            print(f"  错误: 文件 {input_json_path} 未找到。")
        except Exception as e:
            print(f"  错误: 读取或解析文件 {filename} 时发生意外错误: {e}。跳过。")
    
    if json_files_processed == 0 and total_json_files:
        print("\n没有JSON文件被成功处理，请检查错误信息。")
    elif json_files_processed > 0:
        # pylint: disable=line-too-long
        print(f"\n处理完成。共成功处理了 {json_files_processed} / {len(total_json_files)} 个JSON文件。") # noqa E501
    # Case where no JSON files were found is handled
    # at the beginning of the loop


if __name__ == "__main__":
    main()
