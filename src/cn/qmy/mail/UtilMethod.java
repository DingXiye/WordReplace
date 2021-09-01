package cn.qmy.mail;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class UtilMethod {
	/**
	 * 
	 * @param length
	 *            字符串的长度
	 * @param num
	 *            字符串的个数
	 * @return
	 */
	public List<String> genCodes(int length, long num) {
		List<String> results = new ArrayList<String>();
		for (int j = 0; j < num; j++) {
			String val = "";
			Random random = new Random();
			for (int i = 0; i < length; i++) {
				String charOrNum = random.nextInt(2) % 2 == 0 ? "char" : "num"; // 输出字母还是数字
				if ("char".equalsIgnoreCase(charOrNum)){//字母
					int choice = random.nextInt(2) % 2 == 0 ? 65 : 97; // 取得大写字母还是小写字母
					val += (char) (choice + random.nextInt(26));
				} else if ("num".equalsIgnoreCase(charOrNum)){// 数字
					val += String.valueOf(random.nextInt(10));
				}
			}
			val = val.toLowerCase();
			if (results.contains(val)) {
				continue;
			} else {
				results.add(val);
			}
		}
		return results;
	}
}
