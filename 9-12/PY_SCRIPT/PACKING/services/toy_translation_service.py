"""产品名称本地译名；未收录名称要求补充，避免输出虚构翻译。"""
import re

NAMES = {
    "WM小鸭": "WM toy ducks", "迪士尼小鸭": "Disney toy ducks", "搪胶小鸭子": "Vinyl toy ducks",
    "搪胶小鸭": "Vinyl toy ducks", "小鸭子": "Toy ducks", "小鸭": "Toy ducks", "骰子": "Toy dice",
    "股子": "Toy dice", "振动球": "Vibrating toy ball", "硅胶手感球": "Silicone sensory ball",
    "坐姿米奇": "Sitting Mickey Mouse toy", "爬爬彩盒混装": "Assorted crawling toys in color boxes",
    "恐怖小鸭": "Horror-themed toy ducks", "红色公仔": "Red toy figure", "大迪普梅宝": "Large Dipper and Mabel figures",
    "小迪普梅宝比尔猪": "Small Dipper, Mabel, Bill and Waddles figures", "圣诞小鸭": "Christmas toy ducks",
    "爬爬史迪仔": "Crawling Stitch toy", "爬爬米奇": "Crawling Mickey Mouse toy", "爬爬米妮": "Crawling Minnie Mouse toy",
    "米奇": "Mickey Mouse toy", "米妮": "Minnie Mouse toy", "史迪仔": "Stitch toy", "辛巴": "Simba toy",
    "冥王狗": "Pluto toy", "小飞象": "Dumbo toy", "维尼熊": "Winnie the Pooh toy", "玛丽猫": "Marie cat toy",
    "迪士尼大鸭": "Large Disney toy ducks", "超人系列小鸭": "Superman series toy ducks", "情人节小鸭": "Valentine toy ducks",
    "玩具总动员系列": "Toy Story series toys", "漫威系列": "Marvel series toys", "小维尼熊系列": "Winnie the Pooh series toys",
    "经典米奇系列": "Classic Mickey Mouse series toys", "赛车总动员系列": "Cars series toys", "小美人鱼系列": "The Little Mermaid series toys",
    "海洋奇缘系列": "Moana series toys", "大狗": "Large toy dog", "大马": "Large toy horse",
    "复活节小鸭": "Easter toy ducks", "史迪仔复活节小鸭": "Stitch Easter toy ducks", "眼镜鸭系列": "Toy ducks with glasses",
    "14寸大史迪仔": "14-inch Stitch toy", "磁铁史迪仔": "Magnetic Stitch toy", "26850与28670混装": "Assorted 26850 and 28670 toys",
    "迪士尼13款小鸭": "13 assorted Disney toy ducks", "投影米奇": "Mickey Mouse projector toy", "白种人": "Caucasian doll",
    "拉蒂娜": "Latina doll", "AA": "AA", "反派小鸭": "Villain-themed toy ducks", "三款混装": "Three assorted toy styles",
    "小鸭子7款": "Seven assorted toy duck styles", "24款小鸭": "24 assorted toy duck styles", "浮水鸭": "Floating toy ducks",
    "书包狗": "Backpack toy dog", "长颈鹿": "Toy giraffe", "青蛙": "Toy frog", "大象": "Toy elephant", "狮子": "Toy lion",
    "恐龙": "Toy dinosaur", "4款磁铁公仔混装": "Four assorted magnetic toy figures", "小鸭子系列38款混装": "38 assorted toy duck styles",
    "阿焦": "Disgust character toy", "天使": "Angel character toy", "眼镜鸭": "Toy ducks with glasses", "梅根": "Megan toy",
    "钥匙扣公仔混装": "Assorted toy figure keychains", "磁铁公仔（混装）": "Assorted magnetic toy figures",
    "糖果小鸭": "Candy-themed toy ducks", "糖果鸭": "Candy-themed toy ducks", "磁铁公仔": "Magnetic toy figures",
    "六款豆袋系列": "Six assorted beanbag toys", "我爱纽约组合小鸭": "I Love New York toy duck set",
    "封闭盒包装": "Toys in closed-box packaging", "星战系列": "Star Wars series toys", "米妮系列": "Minnie Mouse series toys",
    "史迪仔系列": "Stitch series toys", "蜘蛛侠系列": "Spider-Man series toys", "桶装系列": "Bucket-packed toy series",
}


def translate_toy_name(name):
    text = str(name or "").strip()
    key = re.sub(r"\s+", "", text)
    key = re.sub(r"[（(](?:卡板装|#\d+)[）)]", "", key)
    if key in NAMES:
        return NAMES[key]
    if text and not re.search(r"[\u3400-\u9fff]", text):
        return text
    return ""
