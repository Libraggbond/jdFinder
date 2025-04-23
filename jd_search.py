import pandas as pd
import os
import sys
from playwright.sync_api import sync_playwright
import re
import random
import time

def read_excel_data(file_path):
    """
    从Excel文件中读取品牌和商品名称两列的内容
    
    参数:
        file_path (str): Excel文件的路径
    
    返回:
        pandas.DataFrame: 包含品牌和商品名称的数据框
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
        
        # 检查是否包含必要的列
        required_columns = ['品牌', '商品名称']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"错误: Excel文件缺少以下列: {', '.join(missing_columns)}")
            print(f"可用的列: {', '.join(df.columns)}")
            return None
        
        # 只保留品牌和商品名称两列
        result_df = df[['品牌', '商品名称']]
        
        # 删除空值行
        result_df = result_df.dropna()
        
        print(f"成功读取 {len(result_df)} 条数据")
        return result_df
    
    except Exception as e:
        print(f"读取Excel文件时发生错误: {str(e)}")
        return None

def search_jd_with_product(product_name, brand_name, browser, page, results):
    """
    使用Playwright搜索指定商品名称，并从结果页面中提取包含"旗舰"的店铺名称、价格和链接
    
    参数:
        product_name (str): 要搜索的商品名称
        brand_name (str): 商品品牌
        browser: Playwright浏览器实例
        page: Playwright页面实例
        results: 存储结果的列表
    
    返回:
        bool: 搜索是否成功
    """
    try:
        print(f"\n正在搜索商品: {product_name}")
        
        # 等待搜索框加载完成
        print("等待搜索框加载...")
        page.wait_for_selector("#key")
        
        # 清空搜索框
        page.fill("#key", "")
        
        # 在搜索框中输入商品名称
        print(f"在搜索框中输入商品名称: {product_name}")
        page.fill("#key", product_name)
        
        # 按回车键执行搜索
        print("按回车键执行搜索...")
        page.press("#key", "Enter")
        
        # 等待搜索结果加载，增加超时时间到60秒
        print("等待搜索结果加载 (networkidle)...")
        page.wait_for_load_state("networkidle", timeout=60000) # 增加 timeout 参数
    
        # !! 新增：显式等待商品列表容器加载完成 !!
        print("等待商品列表容器 (#J_goodsList) 加载...")
        page.wait_for_selector("#J_goodsList ul.gl-warp > li", timeout=60000) # 等待第一个商品项出现
        
        # 获取当前URL
        current_url = page.url
        print(f"当前页面URL: {current_url}")
        
        # 检查是否跳转到风险验证页面
        if "cfe.m.jd.com/privatedomain/risk_handler" in current_url:
            print("检测到风险验证页面，请完成验证...")
            # 等待用户完成验证，验证完成后会跳转到搜索结果页面
            # 等待URL变化，不再是风险验证页面
            page.wait_for_function(
                """() => !window.location.href.includes('cfe.m.jd.com/privatedomain/risk_handler')""",
                timeout=300000  # 5分钟超时
            )
            print("验证完成，已跳转到搜索结果页面")
            
            # 再次等待页面加载完成，增加超时时间到60秒
            print("再次等待页面加载完成 (networkidle)...")
            page.wait_for_load_state("networkidle", timeout=60000) # 增加 timeout 参数
        
        # !! 新增：验证后再次显式等待商品列表容器加载完成 !!
        print("再次等待商品列表容器 (#J_goodsList) 加载...")
        page.wait_for_selector("#J_goodsList ul.gl-warp > li", timeout=60000) # 等待第一个商品项出现
        
        # !! 新增：增加短暂延时确保页面稳定 !!
        print("增加短暂延时 (1秒)...")
        page.wait_for_timeout(1000) # 1秒延时

        # 获取当前URL (搜索结果页面)
        result_url = page.url
        print(f"搜索完成，最终结果页面URL: {result_url}") # 修改打印信息以区分

        # !! 新增：检查最终URL是否看起来像搜索结果页 (包含 search.jd.com) !!
        if "search.jd.com" not in result_url:
             print(f"警告: 最终URL '{result_url}' 可能不是预期的搜索结果页面。")
             # 可以选择在这里添加更详细的错误处理或日志记录

        # --- 新增：点击销量按钮并等待页面更新 ---
        try:
            print("正在点击'销量'按钮进行排序...")
            # 使用更精确的定位器定位包含“销量”文本的链接
            sales_button_selector = "div.f-sort a:has-text('销量')" 
            page.wait_for_selector(sales_button_selector, timeout=10000) # 等待按钮出现
            page.click(sales_button_selector)
            print("'销量'按钮已点击")

            # 点击后等待网络空闲，让页面重新加载排序后的结果
            print("等待销量排序结果加载 (networkidle)...")
            page.wait_for_load_state("networkidle", timeout=60000)
            # 再次等待商品列表出现，确保排序完成
            print("再次等待商品列表容器 (#J_goodsList) 加载...")
            page.wait_for_selector("#J_goodsList ul.gl-warp > li", timeout=60000) 
            print("销量排序完成，页面已更新")
            # 短暂延时确保渲染完成
            page.wait_for_timeout(1000) 

        except Exception as sort_error:
            print(f"点击'销量'按钮或等待排序结果时出错: {sort_error}")
            # 这里可以选择是继续尝试抓取还是标记为失败，目前选择继续
        # --- 新增结束 ---


        # --- 修改开始：遍历商品项，只提取旗舰店信息 ---
        print("正在查找商品列表项...")
        # 定位到每个商品项的容器，通常是 li 元素
        product_items = page.query_selector_all("#J_goodsList ul.gl-warp > li.gl-item") 
        print(f"找到 {len(product_items)} 个商品项")

        found_flagship = False
        for item in product_items:
            shop_name = ""
            price_value = ""
            price_text = ""
            product_link = ""
            extracted_title = "" # !! 新增：用于存储提取的标题

            # 尝试在当前商品项内查找店铺名称
            shop_element = item.query_selector("a.curr-shop.hd-shopname")
            if shop_element:
                shop_name = shop_element.get_attribute("title") or "" # 获取 title 属性作为店铺名

            # 如果店铺名称包含 "旗舰"，则继续提取价格、链接和标题
            if shop_name and "旗舰" in shop_name:
                found_flagship = True
                print(f"找到旗舰店铺: {shop_name}")

                # 尝试在当前商品项内查找价格
                price_element = item.query_selector("div.p-price i[data-price]") # 更精确地定位价格元素
                if price_element:
                    price_value = price_element.get_attribute("data-price") or ""
                    price_text = price_element.inner_text() or ""
                    print(f"  价格: {price_text}元 (原始值: {price_value})")
                else:
                     print("  未找到价格信息")


                # 尝试在当前商品项内查找商品链接
                link_element = item.query_selector("div.p-img > a[href]")
                if link_element:
                    href = link_element.get_attribute("href")
                    if href and href.startswith("//"):
                        product_link = "https:" + href
                        print(f"  链接: {product_link}")
                    elif href:
                         product_link = href # 如果不是 // 开头，直接使用
                         print(f"  链接: {product_link}")
                else:
                    print("  未找到商品链接")

                # !! 新增：尝试在当前商品项内查找并提取商品标题 !!
                title_em_element = item.query_selector("div.p-name a em") # 定位到包含标题的<em>标签
                if title_em_element:
                    extracted_title = title_em_element.inner_text().strip() # 获取纯文本并去除首尾空格
                    print(f"  提取的标题: {extracted_title}")
                else:
                    print("  未找到商品标题元素 (div.p-name a em)")


                # 将找到的旗舰店信息添加到结果列表
                result_item = {
                    "品牌": brand_name,
                    "商品名称": product_name, # 这是Excel输入的原始商品名
                    "旗舰店铺": shop_name,
                    "价格值": price_value,
                    "显示价格": price_text,
                    "商品链接": product_link,
                    "提取的商品标题": extracted_title # !! 新增：添加提取的标题
                }
                results.append(result_item)
                print("-" * 30) # 分隔每个找到的旗舰店信息

        if not found_flagship:
            print(f"商品: {product_name} 未找到符合条件的旗舰店铺")
            # 可以选择是否为未找到旗舰店的商品添加一条空记录
            # results.append({
            #     "品牌": brand_name,
            #     "商品名称": product_name,
            #     "旗舰店铺": "未找到",
            #     "价格值": "",
            #     "显示价格": "",
            #     "商品链接": ""
            # })

        # --- 修改结束 ---

        return True
        
    except Exception as e:
        print(f"搜索商品 '{product_name}' 时发生错误: {str(e)}") # 在错误信息中包含商品名
        # 即使发生错误，也添加一条记录
        results.append({
            "品牌": brand_name,
            "商品名称": product_name,
            "旗舰店铺": "搜索失败",
            "价格值": "",
            "显示价格": "",
            "商品链接": f"错误: {str(e)}",
            "提取的商品标题": "" # !! 新增：错误时也添加空标题列
        })
        return False

def main():
    # 检查命令行参数
    if len(sys.argv) < 2:
        print("使用方法: python jd_search.py <Excel文件路径>")
        return
    
    # 从命令行参数获取文件路径
    file_path = sys.argv[1]
    
    # 读取数据
    data = read_excel_data(file_path)
    
    if data is None or data.empty:
        print("无法从Excel文件中获取商品数据")
        return
    
    print(f"准备搜索 {len(data)} 个商品")
    
    # 存储所有搜索结果
    all_results = []
    
    with sync_playwright() as p:
        # 启动浏览器
        browser = p.chromium.launch(headless=False)  # headless=False 可以看到浏览器界面
        
        try:
            # 创建新页面
            page = browser.new_page()
            
            # 先访问京东登录页面
            login_url = "https://passport.jd.com/new/login.aspx?ReturnUrl=https%3A%2F%2Fwww.jd.com%2F"
            print(f"正在访问京东登录页面: {login_url}")
            page.goto(login_url)
            
            # 等待用户手动登录
            print("请在浏览器中完成登录操作...")
            
            # 等待登录完成，检测是否跳转到京东首页
            page.wait_for_url("https://www.jd.com/**", timeout=300000)  # 设置5分钟超时，等待用户登录
            print("登录成功，已跳转到京东首页")
            
            # 遍历所有商品进行搜索
            for index, row in data.iterrows():
                product_name = row['商品名称']
                brand_name = row['品牌']
                print(f"\n===== 正在处理第 {index+1}/{len(data)} 个商品 =====")
                
                # 搜索商品
                success = search_jd_with_product(product_name, brand_name, browser, page, all_results)
                
                if not success:
                    print(f"搜索商品 '{product_name}' 失败，已记录")
                
                # 每个商品搜索后暂停一下，随机等待5-10秒
                if index < len(data) - 1:  # 如果不是最后一个商品
                    wait_time = random.uniform(5, 10)
                    print(f"\n请查看当前商品的搜索结果，{wait_time:.1f}秒后将继续搜索下一个商品...")
                    page.wait_for_timeout(int(wait_time * 1000))  # 转换为毫秒
            
            # 所有商品搜索完成，保存结果到Excel
            result_file = "result.xlsx"
            print(f"\n所有商品搜索完成，正在保存结果到 {result_file}...")
            
            # 创建DataFrame并保存
            result_df = pd.DataFrame(all_results)
            result_df.to_excel(result_file, index=False)
            
            print(f"结果已保存到 {result_file}")
            print("按Ctrl+C终止程序...")
            
            # 等待用户手动终止程序
            page.wait_for_timeout(60000)  # 等待1分钟
            
        except Exception as e:
            print(f"发生错误: {str(e)}")
            
            # 如果已经有搜索结果，尝试保存
            if all_results:
                try:
                    result_file = "result.xlsx"
                    print(f"尝试保存已有结果到 {result_file}...")
                    result_df = pd.DataFrame(all_results)
                    result_df.to_excel(result_file, index=False)
                    print(f"结果已保存到 {result_file}")
                except Exception as save_error:
                    print(f"保存结果时发生错误: {str(save_error)}")
        finally:
            # 关闭浏览器
            browser.close()
            print("浏览器已关闭")

if __name__ == "__main__":
    main()