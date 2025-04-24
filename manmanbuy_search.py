import pandas as pd
import os
import sys
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import random
import time
import re # <-- 新增：导入 re 模块
from urllib.parse import unquote # <-- 新增：用于解码 URL

def read_excel_data_manmanbuy(file_path):
    """
    从Excel文件中读取商品名称列的内容
    
    参数:
        file_path (str): Excel文件的路径
    
    返回:
        pandas.DataFrame: 包含商品名称的数据框，列名为 '商品名称'
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件 '{file_path}' 不存在")
            return None
        
        # 检查文件扩展名
        if not file_path.endswith(('.xlsx', '.xls')):
            print(f"错误: 文件 '{file_path}' 不是Excel文件")
            return None
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 检查是否包含 '商品名称' 列
        required_column = '商品名称'
        if required_column not in df.columns:
            print(f"错误: Excel文件缺少列: '{required_column}'")
            print(f"可用的列: {', '.join(df.columns)}")
            return None
        
        # 只保留商品名称列
        result_df = df[[required_column]]
        
        # 删除空值行
        result_df = result_df.dropna()
        
        print(f"成功读取 {len(result_df)} 条商品名称数据")
        return result_df
    
    except Exception as e:
        print(f"读取Excel文件时发生错误: {str(e)}")
        return None


# !! 修改：函数现在返回提取到的数据列表 !!
def search_manmanbuy_product(product_name, page):
    """
    在慢慢买网站上搜索指定的商品名称, 并提取结果中的商品名、链接、价格、平台和店铺
    
    参数:
        product_name (str): 要搜索的商品名称
        page: Playwright页面实例
    
    返回:
        list: 包含提取到的商品信息的字典列表
    """
    extracted_items = []
    try:
        print(f"\n正在搜索商品: {product_name}")
        
        search_box_selector = "#skey"
        print(f"等待搜索框 '{search_box_selector}' 加载...")
        page.wait_for_selector(search_box_selector, timeout=10000)
        
        print(f"在搜索框中输入商品名称: {product_name}")
        page.fill(search_box_selector, product_name)
        
        print("按回车键执行搜索...")
        page.press(search_box_selector, "Enter")

        # 等待搜索结果加载
        print("等待搜索请求后的网络活动稳定 (networkidle)...")
        try:
            page.wait_for_load_state("networkidle", timeout=30000) 
            print("网络活动已稳定。")
            print("短暂延时 (2秒) 确保内容渲染...")
            page.wait_for_timeout(2000)
        except PlaywrightTimeoutError:
            print("警告: 等待 networkidle 超时，可能仍在加载或已加载完成。继续尝试查找结果...")

        # 查找所有商品div
        product_divs = page.query_selector_all("div.bjlineSmall")
        print(f"找到 {len(product_divs)} 个商品项")

        # 遍历每个商品div
        for i, product_div in enumerate(product_divs):
            try:
                print(f"\n正在处理第 {i + 1}/{len(product_divs)} 个商品...")
                
                # 获取商品div的完整HTML
                div_html = product_div.evaluate("node => node.outerHTML")
                
                # 1. 从onclick属性提取商品名称和价格
                onclick_match = re.search(r"uploadEvent\('([^']*)','\d+','[^']*','[^']*','[^']*','(\d+(?:\.\d+)?)'", div_html)
                if not onclick_match:
                    print("  未找到商品名称和价格信息，跳过此商品")
                    continue
                    
                extracted_name = onclick_match.group(1).strip()
                extracted_price = onclick_match.group(2).strip()
                
                # 2. 提取商品链接
                extracted_url = ""  # 默认为空字符串
                url_match = re.search(r'originalUrl=([^&"]+)', div_html)
                if url_match:
                    raw_url = unquote(url_match.group(1))
                    # 修改：增加对京粉链接的支持
                    if raw_url.startswith("https://item.jd.com"):
                        extracted_url = raw_url
                    elif raw_url.startswith("http://item.jd.com"):
                        extracted_url = "https://" + raw_url[len("http://"):]
                    elif raw_url.startswith("https://jingfen.jd.com"):
                        extracted_url = raw_url
                    elif raw_url.startswith("http://jingfen.jd.com"):
                        extracted_url = "https://" + raw_url[len("http://"):]
                else:
                    print("  未找到商品链接，将使用空链接继续处理")

                # 3. 提取平台信息
                platform_matches = re.findall(r'<span\s+class="shenqingGY">\s*([^<]+?)\s*</span>', div_html)
                extracted_platform = platform_matches[0].strip() if platform_matches else ""
                
                # 4. 提取店铺信息
                shop_matches = re.findall(r'<p\s+class="AreaZY">\s*([^<]+?)\s*</p>', div_html)
                extracted_shop = shop_matches[0].strip() if shop_matches else ""
                
                # 5. 如果所有必要信息都已提取到，则添加到结果列表
                if extracted_name and extracted_url and extracted_price:
                    print(f"  提取到: 名称='{extracted_name}', 价格='{extracted_price}', "
                          f"平台='{extracted_platform}', 店铺='{extracted_shop}', 链接='{extracted_url}'")
                    
                    extracted_items.append({
                        "name": extracted_name,
                        "price": extracted_price,
                        "platform": extracted_platform,
                        "shop": extracted_shop,
                        "url": extracted_url
                    })
                
            except Exception as item_error:
                print(f"  处理商品项时出错: {item_error}")
                continue

        print(f"商品 '{product_name}' 搜索完成，共提取到 {len(extracted_items)} 条有效结果。")
        return extracted_items

    except PlaywrightTimeoutError as te:
        print(f"错误: 搜索商品 '{product_name}' 时发生超时: {te}")
        return extracted_items
    except Exception as e:
        print(f"搜索商品 '{product_name}' 时发生错误: {str(e)}")
        return extracted_items


def main():
    # 检查命令行参数
    if len(sys.argv) < 2:
        print("使用方法: python manmanbuy_search.py <Excel文件路径>")
        return
    
    file_path = sys.argv[1]
    data = read_excel_data_manmanbuy(file_path)
    
    if data is None or data.empty:
        print("无法从Excel文件中获取商品数据")
        return
    
    print(f"准备搜索 {len(data)} 个商品")

    all_results = [] 

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False) 
        page = browser.new_page()
        
        try:
            # 访问慢慢买首页
            home_url = "http://www.manmanbuy.com/"
            print(f"正在访问慢慢买首页: {home_url}")
            page.goto(home_url, wait_until="networkidle") 

            login_button_selector = "a.pt[onclick*='loginShow']" 
            try:
                print(f"查找并点击登录按钮: {login_button_selector}")
                page.wait_for_selector(login_button_selector, timeout=15000)
                page.click(login_button_selector)
                print("登录按钮已点击，请在浏览器中完成登录操作...")
                
                print("等待登录完成（检测 body style.overflow 变为 'auto'）...")
                try:
                    page.wait_for_function(
                        """() => document.body.style.overflow === 'auto'""",
                        timeout=300000  
                    )
                    print("检测到登录完成（body style.overflow 已变为 'auto'）！")
                except PlaywrightTimeoutError:
                    print("等待登录超时（5分钟），将继续执行...")
                except Exception as wait_error:
                     print(f"等待登录状态变化时发生错误: {wait_error}")
                     print("将继续执行...")

            except PlaywrightTimeoutError:
                 print("错误：未能找到登录按钮，请检查页面结构或选择器。")


            # 遍历所有商品进行搜索
            for index, row in data.iterrows():
                product_name = row['商品名称']
                print(f"\n===== 正在处理第 {index+1}/{len(data)} 个商品: {product_name} =====")
                
                extracted_data = search_manmanbuy_product(product_name, page)
                
                if extracted_data:
                    for item in extracted_data:
                        # !! 修改：添加平台和店铺到结果字典 !!
                        all_results.append({
                            "搜索词": product_name,
                            "提取的商品名": item["name"],
                            "价格": item["price"], 
                            "平台": item["platform"], # <-- 新增
                            "店铺": item["shop"],     # <-- 新增
                            "商品链接": item["url"]
                        })
                else:
                    # !! 修改：未找到结果时也添加平台和店铺列 !!
                    all_results.append({
                        "搜索词": product_name,
                        "提取的商品名": "未找到匹配结果",
                        "价格": "", 
                        "平台": "", # <-- 新增
                        "店铺": "", # <-- 新增
                        "商品链接": ""
                    })
                
                # 每个商品搜索后暂停一下，随机等待3-7秒
                if index < len(data) - 1:
                    wait_time = random.uniform(3, 7)
                    print(f"\n处理完成，暂停 {wait_time:.1f} 秒后继续...")
                    page.wait_for_timeout(int(wait_time * 1000))
            
            print("\n所有商品处理完成。")
            
            # !! 修改：对结果进行去重 (包含平台和店铺) !!
            print(f"\n原始结果数量: {len(all_results)}")
            unique_results = []
            seen_combinations = set() # 用于存储已经见过的组合

            for item in all_results:
                if item["提取的商品名"] != "未找到匹配结果":
                    # !! 修改：组合键包含平台和店铺 !!
                    combination_key = (
                        item["提取的商品名"], 
                        item["价格"], 
                        item["平台"], 
                        item["店铺"], 
                        item["商品链接"]
                    )
                    
                    if combination_key not in seen_combinations:
                        unique_results.append(item)
                        seen_combinations.add(combination_key)
                else:
                    unique_results.append(item)
            
            print(f"去重后结果数量: {len(unique_results)}")
            # !! 去重结束 !!

            # !! 修改：使用去重后的 unique_results 保存到 Excel (列名已更新) !!
            if unique_results: 
                print("\n正在将去重后的结果保存到 Excel 文件...")
                # !! 修改：更新列顺序 !!
                columns_order = ["搜索词", "提取的商品名", "价格", "平台", "店铺", "商品链接"] 
                results_df = pd.DataFrame(unique_results, columns=columns_order) 
                output_filename = "manmanbuy_results.xlsx"
                try:
                    results_df.to_excel(output_filename, index=False)
                    print(f"结果已成功保存到: {output_filename}")
                except Exception as save_error:
                    print(f"保存结果到 Excel 时出错: {save_error}")
            else:
                print("没有提取到任何结果或所有结果都被去重，未生成 Excel 文件。")

            print("脚本将在10秒后自动关闭，您可以手动关闭浏览器。")
            page.wait_for_timeout(10000) 

        except Exception as e:
            print(f"\n在主流程中发生错误: {str(e)}")
            # !! 修改：出错时也尝试保存去重后的部分结果 (包含平台和店铺) !!
            print("\n尝试对已收集的结果进行去重...")
            unique_partial_results = []
            seen_partial_combinations = set()
            for item in all_results: 
                 if item["提取的商品名"] != "未找到匹配结果":
                     # !! 修改：组合键包含平台和店铺 !!
                     combination_key = (
                         item["提取的商品名"], 
                         item["价格"], 
                         item["平台"], 
                         item["店铺"], 
                         item["商品链接"]
                     )
                     if combination_key not in seen_partial_combinations:
                         unique_partial_results.append(item)
                         seen_partial_combinations.add(combination_key)
                 else:
                     unique_partial_results.append(item)
            
            if unique_partial_results:
                 print(f"去重后部分结果数量: {len(unique_partial_results)}")
                 print("尝试保存部分结果到 Excel 文件...")
                 # !! 修改：更新列顺序 !!
                 columns_order = ["搜索词", "提取的商品名", "价格", "平台", "店铺", "商品链接"]
                 results_df = pd.DataFrame(unique_partial_results, columns=columns_order) 
                 output_filename = "manmanbuy_partial_results.xlsx"
                 try:
                     results_df.to_excel(output_filename, index=False)
                     print(f"部分结果已成功保存到: {output_filename}")
                 except Exception as save_error:
                     print(f"保存部分结果到 Excel 时出错: {save_error}")

        finally:
            print("正在关闭浏览器...")
            browser.close()
            print("浏览器已关闭")

if __name__ == "__main__":
    main()
