package kaitou.office.excel.common;

/**
 * 系统代码.
 * User: 赵立伟
 * Date: 2015/1/10
 * Time: 22:00
 */
public class SysCode {
    public enum CellType {
        STRING(1), DATE(2), NUMERIC(3);

        private int value;

        CellType(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    public static final int CELL_TYPE_STRING = 1;
    public static final int CELL_TYPE_DATE = 2;
    public static final int CELL_TYPE_NUMERIC = 3;

    public enum Code {
        P("WFP", "TDS", "M"), T("WFP", "TDS", "M"), F("WFP", "TDS", "M"), A("WFP", "DGS", "M"), V("CPP", "DP", ""), C("WFP", "TDS", ""), I("CPP", "PGA", "");
        private String cardNoPref;
        private String models;
        private String readUnit;

        Code(String cardNoPref, String models, String readUnit) {
            this.cardNoPref = cardNoPref;
            this.models = models;
            this.readUnit = readUnit;
        }

        public String getCardNoPref() {
            return cardNoPref;
        }

        public String getModels() {
            return models;
        }

        public String getReadUnit() {
            return readUnit;
        }

        public static Code getCode(String value) {
            for (Code code : Code.values()) {
                if (code.name().equals(value)) {
                    return code;
                }
            }
            return P;
        }
    }
}
