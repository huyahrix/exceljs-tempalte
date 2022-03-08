const ERROR_TYPE_UNKNOWN = 1;
const ERROR_TYPE_INPUT = 2;
const ERROR_TYPE_PARSE = 3;
// export const ERROR_VERIFY_SIGNATURE = 4;

class ExceljsTemplateError extends Error {
    constructor(msg, type = ERROR_TYPE_UNKNOWN) {
        super(msg);
        this.type = type;
    }
}

// Shorthand
ExceljsTemplateError.TYPE_UNKNOWN = ERROR_TYPE_UNKNOWN;
ExceljsTemplateError.TYPE_INPUT = ERROR_TYPE_INPUT;
ExceljsTemplateError.TYPE_PARSE = ERROR_TYPE_PARSE;
// SignPdfError.VERIFY_SIGNATURE = ERROR_VERIFY_SIGNATURE;

module.exports = ExceljsTemplateError;
