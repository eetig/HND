// 数据源
const dataList = [
    { a: 'HND-D171(1.2-双三甲氧基硅基乙烷)', b: '200/塑桶', c: 115055982 },
    { a: 'HND-V171(乙烯基三甲氧基硅烷)', b: '950/吨桶', c: 115056207 },
    { a: 'HND-V171(乙烯基三甲氧基硅烷)', b: '190/铁桶', c: 115055989 },
    { a: 'HND-V171(乙烯基三甲氧基硅烷)', b: '28/塑桶', c: 115058270 },
    { a: 'HND-V171(乙烯基三甲氧基硅烷)', b: '5/塑桶', c: 115056204 },
    { a: 'HND-V171(乙烯基三甲氧基硅烷)', b: '吨桶', c: 115056207 },
    { a: 'HND-V171(乙烯基三甲氧基硅烷)', b: '25/塑桶', c: 115056205 },
    { a: 'HND-V150(乙烯基三氯硅烷)', b: '220/塑桶', c: 115055980 },
    { a: 'HND-V150(乙烯基三氯硅烷)', b: '槽罐', c: 115057979 },
    { a: 'HND-TMOS(L)(正硅酸甲酯)', b: '吨桶', c: 115056001 },
    { a: 'HND-TMOS(L)(正硅酸甲酯)', b: '190/塑桶', c: 115056301 },
    { a: 'HND-TMOS(L)(正硅酸甲酯)', b: '200/铁桶', c: 115056302 },
    { a: 'HND-TMOS(L)(正硅酸甲酯)', b: '200/塑桶', c: 115056301 },
    { a: '电石渣', b: '散装', c: 114001903 },
    { a: 'HND-D150(1.2-双三氯硅基乙烷)', b: '250/铁桶', c: 115055993 },
    { a: 'HND-D150(1.2-双三氯硅基乙烷)', b: '200L塑桶', c: 115055983 },
    { a: 'HND-D150(1.2-双三氯硅基乙烷)', b: '200L铁桶', c: 115055993 }
];

// 预构建哈希索引（仅执行1次）
const dataIndex = {};
for (const item of dataList) {
    // 一级键：a
    if (!dataIndex[item.a]) {
        dataIndex[item.a] = {};
    }
    // 二级键：b，值：c
    dataIndex[item.a][item.b] = item.c;
}




/**
 * 171精馏：产品A/B/C罐
 * 0:不算封头
 * 1:算封头
 * @customfunction
 */
function V171(h,f) {
    if (f==0) {
        return h*0.00314*1000
    }else if (f==1) {
        return (h*0.00314)*1000+1596.7
    }
}

/**
 * 计算卧式带有椭圆封头的圆筒形容器中的液体体积
 * r：圆筒半径（直径的一半），椭圆封头长半径与圆筒半径相同
 * l：圆筒长度（不含封头）
 * h_i：椭圆封头短半轴（即椭圆封头内高度）
 * h：液位高度（从底部到液面的高度）
 * @customfunction
 * @param {number} r - 圆筒半径（直径的一半），椭圆封头长半径与圆筒半径相同
 * @param {number} l - 圆筒长度（不含封头）
 * @param {number} h_i - 椭圆封头短半轴（即椭圆封头内高度）
 * @param {number} h - 液位高度（从底部到液面的高度）
 * @returns {number} - 液体体积
 */
function tank_HS(r, l, h_i, h) {
    // 确保所有参数都是数字
    r = Number(r);
    l = Number(l);
    h_i = Number(h_i);
    h = Number(h);
    
    // 检查参数有效性
    if (r <= 0 || l <= 0 || h_i <= 0 || h < 0) {
        return "参数错误：所有参数必须大于0，液位高度不能为负";
    }
    
    // 总高度（直径）
    const D = 2 * r;
    
    // 检查液位高度是否超过总高度
    if (h > D) {
        return "参数错误：液位高度不能超过储罐直径";
    }
    
    // 根据用户提供的公式计算液体体积
    // 公式：V_h = L [ (πr²/2 - (r-h)√(2rh - h²) - r²arcsin((r-h)/r) ] + (πh_i)/(3r) [3r²h - r³ + (r-h)³]
    
    // 1. 计算圆柱部分的液体体积
    const part1 = l * ( (Math.PI * Math.pow(r, 2)) / 2 - (r - h) * Math.sqrt(2 * r * h - Math.pow(h, 2)) - Math.pow(r, 2) * Math.asin( (r - h) / r ) );
    
    // 2. 计算两个椭圆封头的液体体积总和
    const part2 = (Math.PI * h_i) / (3 * r) * ( 3 * Math.pow(r, 2) * h - Math.pow(r, 3) + Math.pow(r - h, 3) );
    
    // 3. 总体积
    const volume = part1 + part2;
    
    return volume;
}

/**
 * 计算立式带有椭圆封头的圆筒形容器中的液体体积
 * r：圆筒半径（直径的一半），椭圆封头长半径与圆筒半径相同
 * l：圆筒高度（不含封头）
 * h_i：椭圆封头短半轴（即椭圆封头内高度）
 * h：液位高度（从底部到液面的高度）
 * @customfunction
 * @param {number} r - 圆筒半径（直径的一半），椭圆封头长半径与圆筒半径相同
 * @param {number} l - 圆筒高度（不含封头）
 * @param {number} h_i - 椭圆封头短半轴（即椭圆封头内高度）
 * @param {number} h - 液位高度（从底部到液面的高度）
 * @returns {number} - 液体体积
 */
function tank_VS(r, l, h_i, h) {
    // 确保所有参数都是数字
    r = Number(r);
    l = Number(l);
    h_i = Number(h_i);
    h = Number(h);
    
    // 检查参数有效性
    if (r <= 0 || l <= 0 || h_i <= 0 || h < 0) {
        return "参数错误：所有参数必须大于0，液位高度不能为负";
    }
    
    // 总高度（储罐总高度）
    const total_height = 2 * h_i + l;
    
    // 检查液位高度是否超过总高度
    if (h > total_height) {
        return "参数错误：液位高度不能超过储罐总高度";
    }
    
    let volume = 0;
    
    // 无直边非标椭圆封头的体积公式：V = (2/3)πa²b
    // 这里a为椭圆封头长半径（与圆筒半径相同），b为椭圆封头短半轴
    const V_ellipse_full = (2/3) * Math.PI * Math.pow(r, 2) * h_i;
    
    if (h <= h_i) {
        // 1. 仅底部椭圆封头内有液体
        // 使用用户提供的无直边非标椭圆封头体积公式的液位比例计算
        // 假设液位高度h与封头高度h_i的比例为k，则体积V = (2/3)πr²h_i * (h/h_i)^2
        const k = h / h_i;
        volume = V_ellipse_full * Math.pow(k, 2);
    } else if (h <= h_i + l) {
        // 2. 底部椭圆封头满，圆筒部分有液体
        // 椭圆封头满体积：V_ellipse_full = (2/3)πr²h_i
        // 圆筒部分液体体积：V_cylinder = πr²(h - h_i)
        const V_cylinder = Math.PI * Math.pow(r, 2) * (h - h_i);
        volume = V_ellipse_full + V_cylinder;
    } else {
        // 3. 底部椭圆封头和圆筒都满，顶部椭圆封头内有液体
        // 两个椭圆封头满体积：V_ellipse_full_total = 2 * (2/3)πr²h_i
        const V_ellipse_full_total = 2 * V_ellipse_full;
        // 圆筒满体积：V_cylinder_full = πr²l
        const V_cylinder_full = Math.PI * Math.pow(r, 2) * l;
        // 顶部椭圆封头内的液体体积：V_top_ellipse = (2/3)πr²h_i * (h_top/h_i)^2
        const h_top = (2*h_i + l-h); // 顶部椭圆封头内的气体高度
        const k_top = h_top / h_i;
        const V_top_ellipse = V_ellipse_full * Math.pow(k_top, 2);
        volume = 2*V_ellipse_full + V_cylinder_full - V_top_ellipse;
    }
    
    return volume;
}

/**
 * 根据物料名称和规格匹配返回对应的编码
 * @customfunction
 * @param {string} a - 物料名称
 * @param {string} b - 规格
 * @returns {string} - 匹配的编码
 */
function wlbm(a, b) {
    return dataIndex[a]?.[b] ?? "";
}


/**
 * 根据液位计算三氯氢硅的体积
 * @customfunction
 * @param {string} h - 三氯氢硅液位（单位mm）
 */
function V_SiHCl3(h) {
    return tank_HS(1.4, 5.58, 0.7, h/1000);
}


/**
 * 根据液位计算三氯氢硅的体积
 * π=3.14
 * @customfunction
 * @param {string} h - 三氯氢硅液位（单位mm）
 */
function V_SiHCl3_314(h) {
    // 确保所有参数都是数字
    const r = 1.4;
    const l = 5.58;
    const h_i = 0.7;
    h = Number(h);
    
    // 检查参数有效性
    if (r <= 0 || l <= 0 || h_i <= 0 || h < 0) {
        return "参数错误：所有参数必须大于0，液位高度不能为负";
    }
    
    // 总高度（直径）
    const D = 2 * r;
    
    // 检查液位高度是否超过总高度
    if (h > D) {
        return "参数错误：液位高度不能超过储罐直径";
    }
    
    // 根据用户提供的公式计算液体体积
    // 公式：V_h = L [ (πr²/2 - (r-h)√(2rh - h²) - r²arcsin((r-h)/r) ] + (πh_i)/(3r) [3r²h - r³ + (r-h)³]
    
    // 1. 计算圆柱部分的液体体积
    const part1 = l * ( (3.14 * Math.pow(r, 2)) / 2 - (r - h) * Math.sqrt(2 * r * h - Math.pow(h, 2)) - Math.pow(r, 2) * Math.asin( (r - h) / r ) );
    
    // 2. 计算两个椭圆封头的液体体积总和
    const part2 = (3.14 * h_i) / (3 * r) * ( 3 * Math.pow(r, 2) * h - Math.pow(r, 3) + Math.pow(r - h, 3) );
    
    // 3. 总体积
    const volume = part1 + part2;
    
    return volume;
}