% GitHub: github.com/Oslomayor/Pixels-in-Excel
% 2018-8-9
% 返回数字列对应的 Excel 字母列
% [col] = num2col_recur(num)
% 1 2 3...26 27... -> A B C... Z AA...
% Example:
% col = num2col(28) 
% Output: col = AB

function col = num2col_recur(num)
    s= '';
    while num > 0
        s = [char(mod((num - 1),26)+65),s];
        num = fix((num - 1)/26);
    end
    col = s;
end
