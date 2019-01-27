
enum Digits
{
    Units,
    Tens,
    Hundreds,
    Thousands,
    Higher
}

// класс для перевода целого числа в строковую запись

public class Number {

    // число
    private String num;
    // строковое представление
    private String numStr;
    // класс по наибольшей степени
    private Digits numClass;

    private boolean negative = false;

    private static String digitWords[] = { "миллион", "миллиард", "биллион", "триллион", "триллиард", "квадриллион",
            "квадриллиард", "квантиллион", "квантиллиард", "секстиллион", "секстиллиард", "септиллион", "септиллиард",
            "октиллион", "октиллиард", "нониллион", "дециллион", "дециллиард", "ундециллион", "ундециллиард",
            "додециллион", "додециллиард", "тредециллион", "тредециллиард", "кваттуордециллион", "кваттуордециллиард",
            "квиндециллион", "квиндециллиард", "седециллион", "седециллиард", "септдециллион", "септдециллиард",
            "октодециллион", "октодециллиард", "новемдециллион", "новемдециллиард", "вигинтиллион", "вигинтиллиард",
            "анвигинтиллион", "анвигинтиллиард", "дуовигинтиллион", "дуовигинтиллиард", "тревигинтиллион",
            " тревигинтиллиард", "кватторвигинтиллион", "кватторвигинтиллиард", "квинвигинтиллион", "квинвигинтиллиард",
            "сексвигинтиллион", "сексвигинтиллиард", "септемвигинтиллион", "септемвигинтиллиард", "октовигинтиллион",
            "октовигинтиллиард", "новемвигинтиллион", "новемвигинтиллиард", "тригинтиллион", "тригинтиллиард", "" +
            "антригинтиллион", "антригинтиллиард", "дуотригинтиллион", "дуотригинтиллиард"};

    public Number(String val) throws Exception {

        String regex = "\\d+";
        if ((!val.matches(regex) && !val.startsWith("-")) || (val.startsWith("-") && !val.substring(1).matches(regex)) || val.isEmpty() || (val.startsWith("0") && val.length() > 1)) throw new Exception("Неверный формат введенных данных!\t" + val);
        if (val.length() > 192) throw new Exception("Слишком большое число!\t" + val);

        if(val.startsWith("-")){
            num = val.substring(1);
            negative = true;
        }
        else
            num = val;

        switch(num.length()) {
            case 1:
                numClass = Digits.Units;
                break;
            case 2:
                numClass = Digits.Tens;
                break;
            case 3:
                numClass = Digits.Hundreds;
                break;
            case 4:
                numClass = Digits.Thousands;
                break;
            case 5:
                numClass = Digits.Thousands;
                break;
            default:
                numClass = Digits.Higher;
                break;
        }
    }

    public String getString() {
        if (!numStr.isEmpty())
            return numStr;
        else{
            toString();
            return numStr;
        }
    }

    public String getNum() {
        return num;
    }

    public Number getNumber()
    {
        return this;
    }

    @Override
    public String toString()
    {
        switch (numClass)
        {
            case Units:
                numStr = getUnits(Integer.parseInt(num));
                break;
            case Tens:
                numStr = getTens(Integer.parseInt(num));
                break;
            case Hundreds:
                numStr = getHundreds(Integer.parseInt(num));
                break;
            case Thousands:
                numStr = getThousands(Integer.parseInt(num));
                break;
            default:
                numStr = getHigh();
                break;
        }
        if (negative)
            numStr = String.join(" ","минус", numStr);
        return numStr;
    }

    public int getCountOfDigits(int n) {
        int count = (n == 0) ? 1 : 0;
        while (n != 0) {
            count++;
            n /= 10;
        }
        return count;
    }

    public String getUnits(int unit)
    {
        String unitStr;
        switch(unit)
        {
            case 0:
                unitStr = (num.length() == 1) ? "ноль" : "";
                break;
            case 1:
                unitStr = "один";
                break;
            case 2:
                unitStr = "два";
                break;
            case 3:
                unitStr = "три";
                break;
            case 4:
                unitStr = "четыре";
                break;
            case 5:
                unitStr = "пять";
                break;
            case 6:
                unitStr = "шесть";
                break;
            case 7:
                unitStr = "семь";
                break;
            case 8:
                unitStr = "восемь";
                break;
            default:
                unitStr = "девять";
                break;
        }
        return unitStr;
    }

    public String getTens(int ten)
    {
        String tenStr;
        if (ten == 10)
            tenStr = "десять";
        else if(ten < 20)
        {
            switch(ten)
            {
                case 11:
                    tenStr = "один";
                    break;
                case 12:
                    tenStr = "две";
                    break;
                case 13:
                    tenStr = "три";
                    break;
                case 14:
                    tenStr = "четыр";
                    break;
                case 15:
                    tenStr = "пят";
                    break;
                case 16:
                    tenStr = "шест";
                    break;
                case 17:
                    tenStr = "сем";
                    break;
                case 18:
                    tenStr = "восем";
                    break;
                default:
                    tenStr = "девят";
                    break;
            }
            tenStr = tenStr.concat("надцать");
        }
        else
        {
            int n = ten/10;
            switch(n)
            {
                case 2:
                    tenStr = "двадцать";
                    break;
                case 3:
                    tenStr = "тридцать";
                    break;
                case 4:
                    tenStr = "сорок";
                    break;
                case 5:
                    tenStr = "пять";
                    break;
                case 6:
                    tenStr = "шесть";
                    break;
                case 7:
                    tenStr = "семь";
                    break;
                case 8:
                    tenStr = "восемь";
                    break;
                default:
                    tenStr = "девяносто";
                    break;
            }
            if (n >=5 && n<=8)
                tenStr = tenStr.concat("десят");
            if(ten%10 !=0)
                tenStr = String.join(" ", tenStr, getUnits(ten%10));
        }
        return tenStr;
    }

    public String getHundreds(int hundred)
    {
        String hundredStr;
        switch(hundred/100)
        {
            case 1:
                hundredStr = "сто";
                break;
            case 2:
                hundredStr = "двести";
                break;
            case 3:
                hundredStr = "триста";
                break;
            case 4:
                hundredStr = "четыреста";
                break;
            case 5:
                hundredStr = "пятьсот";
                break;
            case 6:
                hundredStr = "шестьсот";
                break;
            case 7:
                hundredStr = "семьсот";
                break;
            case 8:
                hundredStr = "восемьсот";
                break;
            default:
                hundredStr = "девятьсот";
                break;
        }

        if (hundred%100!=0) {
            if (hundred/10%10 == 0)
                hundredStr = String.join(" ", hundredStr, getUnits(hundred%100));
            else
                hundredStr = String.join(" ", hundredStr, getTens(hundred%100));
        }
        return hundredStr;
    }

    public String getThousands(int thousand)
    {
        String thousandStr;
        int temp = thousand/1000;
        switch(getCountOfDigits(temp))
        {
            case 1:
                switch (temp) {
                    case 1:
                        thousandStr = "тысяча";
                        break;
                    case 2:
                        thousandStr = "две тысячи";
                        break;
                    default:
                        thousandStr = getUnits(temp);
                        if(temp <= 4)
                            thousandStr = String.join(" ", thousandStr, "тысячи");
                        else
                            thousandStr = String.join(" ", thousandStr, "тысяч");
                        break;
                }
                break;
            case 2:
                switch (temp%100)
                {
                    case 1:
                        thousandStr = getTens(temp/10*10);
                        thousandStr = String.join(" ", thousandStr, "одна тысяча");
                        break;
                    case 2:
                        thousandStr = getTens(temp/10*10);
                        thousandStr = String.join(" ", thousandStr, "две тысячи");
                        break;
                    default:
                        thousandStr = getTens(temp);
                        break;
                }
                if(temp%10 <= 4 && temp%10 >1)
                    thousandStr = String.join(" ", thousandStr, "тысячи");
                else
                    thousandStr = String.join(" ", thousandStr, "тысяч");
                break;
            default:
                switch (temp%100)
                {
                    case 1:
                        thousandStr = getHundreds(temp/10*10);
                        thousandStr = String.join(" ", thousandStr, "одна тысяча");
                        break;
                    case 2:
                        thousandStr = getHundreds(temp/10*10);
                        thousandStr = String.join(" ", thousandStr, "две тысячи");
                        break;
                    default:
                        thousandStr = getHundreds(temp);
                        if(temp%100 >= 11 && temp%100 < 20)
                            thousandStr = String.join(" ", thousandStr, "тысяч");
                        else if(temp%10 <= 4 && temp%10 >1)
                            thousandStr = String.join(" ", thousandStr, "тысячи");
                        else
                            thousandStr = String.join(" ", thousandStr, "тысяч");
                        break;
                }
                break;
        }
        if(thousand%1000 != 0) {
            if(thousand/10%100 == 0)
                thousandStr += (' ' + getUnits(thousand % 1000));
            else if(thousand/100%10 == 0)
                thousandStr += (' ' + getTens(thousand % 1000));
            else
                thousandStr += (' ' + getHundreds(thousand % 1000));
        }
        return thousandStr;
    }

    public String getHigh()
    {
        String strLow = num.substring(num.length()-6);
        if (!strLow.equals("000000"))
            switch (strLow.length())
            {
                case 1:
                    strLow = getUnits(Integer.parseInt(strLow));
                    break;
                case 2:
                    strLow = getTens(Integer.parseInt(strLow));
                    break;
                case 3:
                    strLow = getHundreds(Integer.parseInt(strLow));
                    break;
                default:
                    strLow = getThousands(Integer.parseInt(strLow));
                    break;

            }
        else
            strLow = "";

        String strHigh = num.substring(0,num.length()-6);
        String newStr = "";
        int temp;
        for(int i = strHigh.length(), degree = 0; i > 0; i -= 3, degree++) {
            if(strHigh.length() <= 3)
                temp = Integer.parseInt(strHigh);
            else {
                temp = Integer.parseInt(strHigh.substring(i-3,i));
                strHigh = strHigh.substring(0,i-3);
            }
            String str;
            switch(getCountOfDigits(temp))
            {
                case 1:
                    str = getUnits(temp);
                    break;
                case 2:
                    str = getTens(temp);
                    break;
                default:
                    str = getHundreds(temp);
                    break;
            }
            if(!str.isEmpty())
                if(temp%100 >= 11 && temp%100 < 20)
                    str = String.join(" ", str, digitWords[degree] + "ов");
                else if (temp%10 == 1)
                    str = String.join(" ", str, digitWords[degree]);
                else if(temp%10 <= 4 && temp%10 >1)
                    str = String.join(" ", str, digitWords[degree] + "а");
                else
                    str = String.join(" ", str, digitWords[degree] + "ов");
            newStr = String.join(" ",str,newStr);
        }
        newStr = newStr.concat(strLow);
        return newStr;
    }

}
