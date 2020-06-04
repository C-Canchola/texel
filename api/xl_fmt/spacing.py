from functools import wraps

from openpyxl.formula.tokenizer import (
    Tokenizer,
    Token
)


def apply_formula_spacing(formula: str) -> str:
    """
    Applies formula spacing to improve readability of excel formulas.
    A wise man once said, flat is better than nested.
    Returns the spaced formula so it can be decided what to do with it.
    Args:
        formula (): str

    Returns:str

    """
    formula = add_arg_and_formula_spacing(formula)
    return space_if_formulas(formula)


def init_formula(func):
    @wraps(func)
    def inner_func(formula, *args, **kwargs):
        formula = remove_white_space(formula)
        return func(formula, *args, **kwargs)

    return inner_func


def get_tokens(cell_formula: str, token_type=None):
    """Returns the tokens from an excel formula.

    Arguments:
        cell_formula {str} -- Excel formula of a cell

    Keyword Arguments:
        token_type {[type]} -- Token type to filter if so desired. (default: {None})

    Returns:
        [type] -- List of tokens.
    """
    items = Tokenizer(cell_formula).items
    if token_type:
        items = [item for item in items if item.subtype == token_type]

    return items


def remove_white_space(formula: str) -> str:
    tokens = get_tokens(formula)
    tkn: Token
    keep_tokens = [tkn for tkn in tokens if tkn.type != Token.WSPACE]

    return convert_tokens_to_string(keep_tokens)


def create_white_space_token() -> Token:
    return Token(' ', Token.WSPACE)


@init_formula
def add_arg_and_formula_spacing(formula: str):
    tokens = get_tokens(formula)
    mk_token_list = []

    tkn: Token
    for tkn in tokens:

        if tkn.subtype in [tkn.ARG]:

            mk_token_list.append(tkn)
            mk_token_list.append(create_white_space_token())

        else:
            mk_token_list.append(tkn)

    return convert_tokens_to_string(mk_token_list)


def convert_tokens_to_string(tokens) -> str:
    first_token: Token = tokens[0]
    pre_append = '=' if first_token.type != Token.LITERAL else ''

    return pre_append + ''.join([tkn.value for tkn in tokens])


def space_if_formulas(formula: str) -> str:
    level = 0
    space_mult = 3
    in_open_non_if_func = False
    in_open_if_func = False
    if_arg_count = 0
    arg_stack = []

    tokens = get_tokens(formula)
    new_tokens = []

    def _token_is_arg_that_needs_space(check_tkn: Token) -> bool:
        """
        Checks if a token is a comma seperator within an if function.

        Args:
            check_tkn (): Token

        Returns: bool

        """
        if not _token_is_comma_arg(check_tkn):
            return False
        if in_open_non_if_func:
            return False
        if not in_open_if_func:
            return False
        return True

    tkn: Token
    for tkn in tokens:

        if not _token_is_open_if(tkn) and if_arg_count == 2:
            if_arg_count = arg_stack.pop()
            level -= 1

        if _token_is_open_if(tkn):
            level += 1
            in_open_if_func = True
            in_open_non_if_func = False
            new_tokens.append(tkn)
            arg_stack.append(if_arg_count)
            if_arg_count = 0

        elif _token_is_arg_that_needs_space(tkn):
            if_arg_count += 1
            new_tokens.append(tkn)

            white_space_str = '\n{}'.format(' ' * level * space_mult)
            white_space_token = Token(white_space_str, Token.WSPACE)
            new_tokens.append(white_space_token)

        elif _token_is_open_func(tkn):
            in_open_non_if_func = True
            new_tokens.append(tkn)

        elif _token_is_closed_func(tkn):
            in_open_non_if_func = False
            new_tokens.append(tkn)
        else:
            new_tokens.append(tkn)

    return convert_tokens_to_string(new_tokens)


def _token_is_open_if(tkn: Token) -> bool:
    if not _token_is_open_func(tkn):
        return False

    if tkn.value != 'IF(':
        return False

    return True


def _token_is_comma_arg(tkn: Token) -> bool:
    if tkn.type != Token.SEP:
        return False
    if tkn.subtype != Token.ARG:
        return False
    if tkn.value != ',':
        return False

    return True


def _token_is_closed_func(tkn: Token) -> bool:
    if tkn.type != Token.FUNC:
        return False
    if tkn.subtype != Token.CLOSE:
        return False
    return True


def _token_is_open_func(tkn: Token) -> bool:
    if tkn.type != Token.FUNC:
        return False

    if tkn.subtype != Token.OPEN:
        return False

    return True
