using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"C:\Users\Computer\OneDrive\Desktop\План 2025\формулы для расчета.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Файл не найден: {filePath}");
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count < 4)
                {
                    Console.WriteLine($"В файле недостаточно листов. Требуется минимум 4 листа, найдено: {package.Workbook.Worksheets.Count}");
                    return;
                }

                var worksheet = package.Workbook.Worksheets[3];
                Console.WriteLine($"Работаем с листом: {worksheet.Name}");

                if (worksheet.Dimension == null)
                {
                    Console.WriteLine("Лист пустой");
                    return;
                }

                int lastRow = worksheet.Dimension.End.Row;
                Console.WriteLine($"Количество строк для обработки: {lastRow}");

                Random random = new Random();

                for (int row = 1; row <= lastRow; row += 2)
                {
                    var targetSum = worksheet.Cells[$"AF{row}"].Value;
                    if (targetSum == null) continue;
                    
                    double targetValue = Convert.ToDouble(targetSum);
                    Console.WriteLine($"Обработка строки {row}, целевая сумма: {targetValue}");
                    
                    if (targetValue % 5 != 0)
                    {
                        Console.WriteLine($"Предупреждение: значение в AF{row} ({targetValue}) не кратно 5");
                        continue;
                    }

                    // Определяем группы ячеек
                    string[][] whiteGroups = new string[][]
                    {
                        new string[] { "C", "D", "E" },  // Группа из 3 ячеек
                        new string[] { },  // Пустая группа - будет пропущена
                        new string[] { "M", "N", "O", "P", "Q" },
                        new string[] { "S", "T", "U", "V", "W" },
                        new string[] { "Y", "Z", "AA", "AB", "AC" }
                    };

                    // Проверяем возможность распределения суммы
                    int nonEmptyGroups = whiteGroups.Count(g => g.Length > 0);
                    double minPossibleSum = nonEmptyGroups * 15; // Минимум 15 на группу
                    double maxPossibleSum = nonEmptyGroups * 200; // Увеличиваем максимум до 200 на группу

                    if (targetValue < minPossibleSum || targetValue > maxPossibleSum)
                    {
                        Console.WriteLine($"Невозможно распределить сумму {targetValue} между {nonEmptyGroups} группами");
                        continue;
                    }

                    bool success = false;
                    int attempts = 0;
                    double[] blueCellValues = new double[5];

                    // Для больших сумм используем другую стратегию
                    bool isLargeSum = targetValue >= nonEmptyGroups * 40;
                    bool isVeryLargeSum = targetValue >= nonEmptyGroups * 80;

                    while (!success && attempts < 5000)
                    {
                        attempts++;
                        success = true;
                        double remainingSum = targetValue;

                        // Распределяем значения по синим ячейкам
                        for (int i = 0; i < whiteGroups.Length; i++)
                        {
                            if (whiteGroups[i].Length == 0)
                            {
                                blueCellValues[i] = 0;
                                continue;
                            }

                            double minBlue = whiteGroups[i].Length * 3;
                            double maxBlue;
                            
                            if (isVeryLargeSum)
                            {
                                maxBlue = 200; // Увеличиваем максимум для очень больших сумм
                            }
                            else if (isLargeSum)
                            {
                                maxBlue = 100;
                            }
                            else
                            {
                                maxBlue = 60;
                            }

                            if (whiteGroups[i].Length <= 3)
                            {
                                maxBlue = Math.Min(maxBlue, whiteGroups[i].Length * 30);
                            }

                            // Округляем до ближайших значений, кратных 5
                            minBlue = Math.Ceiling(minBlue / 5) * 5;
                            maxBlue = Math.Floor(maxBlue / 5) * 5;

                            if (i < whiteGroups.Length - 1)
                            {
                                int remainingGroups = whiteGroups.Skip(i + 1).Count(g => g.Length > 0);
                                double maxPossible = Math.Min(maxBlue, remainingSum - (minBlue * remainingGroups));
                                maxPossible = Math.Floor(maxPossible / 5) * 5;

                                if (maxPossible < minBlue)
                                {
                                    success = false;
                                    break;
                                }

                                if (isVeryLargeSum)
                                {
                                    // Для очень больших сумм используем максимальные значения
                                    double targetForGroup;
                                    if (remainingGroups == 0)
                                    {
                                        targetForGroup = maxPossible;
                                    }
                                    else
                                    {
                                        // Рассчитываем максимально возможное значение для группы
                                        double maxForGroup = whiteGroups[i].Length * 8; // Максимум для белых ячеек
                                        targetForGroup = Math.Min(maxPossible, maxForGroup);
                                        
                                        // Округляем до кратного 5
                                        targetForGroup = Math.Floor(targetForGroup / 5) * 5;
                                        
                                        // Проверяем и корректируем границы
                                        targetForGroup = Math.Min(maxPossible, Math.Max(minBlue, targetForGroup));
                                    }
                                    blueCellValues[i] = targetForGroup;
                                }
                                else if (isLargeSum)
                                {
                                    // Для больших сумм используем более равномерное распределение
                                    double range = maxPossible - minBlue;
                                    double targetForGroup;
                                    
                                    if (remainingGroups == 0)
                                    {
                                        targetForGroup = maxPossible;
                                    }
                                    else
                                    {
                                        double averagePerGroup = remainingSum / (remainingGroups + 1);
                                        targetForGroup = Math.Min(maxPossible, Math.Floor(averagePerGroup / 5) * 5);
                                        
                                        if (range > 10)
                                        {
                                            double randomFactor = random.NextDouble() * 0.2 - 0.1;
                                            targetForGroup = Math.Min(maxPossible, 
                                                Math.Max(minBlue, 
                                                    Math.Floor((targetForGroup * (1 + randomFactor)) / 5) * 5));
                                        }
                                    }
                                    
                                    blueCellValues[i] = targetForGroup;
                                }
                                else
                                {
                                    blueCellValues[i] = Math.Floor(random.NextDouble() * ((maxPossible - minBlue) / 5 + 1)) * 5 + minBlue;
                                }
                                remainingSum -= blueCellValues[i];
                            }
                            else
                            {
                                blueCellValues[i] = remainingSum;
                                if (blueCellValues[i] < minBlue || blueCellValues[i] > maxBlue || blueCellValues[i] % 5 != 0)
                                {
                                    success = false;
                                    break;
                                }
                            }
                        }

                        if (!success) continue;

                        // Для каждой группы белых ячеек
                        for (int i = 0; i < whiteGroups.Length && success; i++)
                        {
                            // Пропускаем пустые группы
                            if (whiteGroups[i].Length == 0)
                            {
                                Console.WriteLine($"Пропускаем пустую группу {i + 1}");
                                continue;
                            }

                            bool groupSuccess = false;
                            int groupAttempts = 0;
                            double[] whiteValues = new double[whiteGroups[i].Length];

                            while (!groupSuccess && groupAttempts < 10)
                            {
                                groupAttempts++;
                                groupSuccess = true;
                                double remainingGroupSum = blueCellValues[i];

                                if (isVeryLargeSum)
                                {
                                    // Для очень больших сумм заполняем максимальными значениями
                                    for (int j = 0; j < whiteGroups[i].Length; j++)
                                    {
                                        whiteValues[j] = 8;
                                    }
                                    groupSuccess = true;
                                }
                                else
                                {
                                    // Генерируем значения для всех ячеек кроме последней
                                    for (int j = 0; j < whiteGroups[i].Length - 1 && groupSuccess; j++)
                                    {
                                        double minValue = Math.Max(3, remainingGroupSum - (8 * (whiteGroups[i].Length - 1 - j)));
                                        double maxValue = Math.Min(8, remainingGroupSum - (3 * (whiteGroups[i].Length - 1 - j)));

                                        if (maxValue < minValue)
                                        {
                                            groupSuccess = false;
                                            break;
                                        }

                                        whiteValues[j] = Math.Floor(minValue + random.NextDouble() * (maxValue - minValue));
                                        remainingGroupSum -= whiteValues[j];
                                    }

                                    if (!groupSuccess) continue;

                                    // Проверяем последнее значение
                                    whiteValues[whiteGroups[i].Length - 1] = Math.Floor(remainingGroupSum);
                                    if (whiteValues[whiteGroups[i].Length - 1] < 3 || whiteValues[whiteGroups[i].Length - 1] > 8)
                                    {
                                        groupSuccess = false;
                                        continue;
                                    }
                                }
                            }

                            if (!groupSuccess)
                            {
                                success = false;
                                break;
                            }

                            // Записываем значения в ячейки
                            for (int j = 0; j < whiteGroups[i].Length; j++)
                            {
                                worksheet.Cells[$"{whiteGroups[i][j]}{row}"].Value = whiteValues[j];
                            }
                        }
                    }

                    if (!success)
                    {
                        Console.WriteLine($"Не удалось сбалансировать значения в строке {row} после {attempts} попыток");
                        continue;
                    }

                    // Устанавливаем формулы для синих ячеек динамически
                    string[] blueCellColumns = new[] { "F", "L", "R", "X", "AD" };
                    for (int i = 0; i < whiteGroups.Length; i++)
                    {
                        if (whiteGroups[i].Length == 0)
                        {
                            worksheet.Cells[$"{blueCellColumns[i]}{row}"].Formula = "=0";
                        }
                        else
                        {
                            string formula = string.Join("+", whiteGroups[i].Select(col => $"{col}{row}"));
                            worksheet.Cells[$"{blueCellColumns[i]}{row}"].Formula = $"={formula}";
                        }
                    }

                    // Формула для AG (сумма всех синих ячеек)
                    string agFormula = string.Join("+", blueCellColumns.Select(col => $"{col}{row}"));
                    worksheet.Cells[$"AG{row}"].Formula = $"={agFormula}";
                    worksheet.Cells[$"AI{row}"].Formula = $"=IF(AND(AF{row}=AG{row}), \"Совпадает\", \"\")";
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.Save();
                Console.WriteLine("Обработка завершена успешно!");

                // Открываем файл Excel
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                    Console.WriteLine("Файл Excel открыт");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Не удалось открыть файл Excel: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
            Console.WriteLine($"Подробности: {ex}");
        }
    }
}
