import xlwings as xw
import texel.naming as txl_nm
import texel.api.name_manager.range_getters as rng_getters
import texel.naming.helpers as nm_helper


class NameStruct:
    """Name struct to help with managing individual names.
    Potential change to namedtuple in the future.
    """

    def __init__(self, nm, refers_to):
        self.nm = nm
        self.refers_to = refers_to

        if self.refers_to[0] == '=':
            self.refers_to = self.refers_to[1:]

    def __hash__(self):
        return hash(self.nm + self.refers_to)

    def __repr__(self):
        return 'NameStuct({}, {})'.format(self.nm, self.refers_to)

    def __eq__(self, other):
        return self.nm == other.nm and self.refers_to == other.refers_to


def get_named_struct_list(nms_and_refs_list):
    """Function to help convert the potential names and references to a single list with
    a structure of two attributes.

    Arguments:
        nms_and_refs_list {list} -- a list of lists, first list names and second refers to.

    Returns:
        list -- list of name structs
    """

    named_struct_list = []
    nms_list = nms_and_refs_list[0:len(nms_and_refs_list):2]
    refs_list = nms_and_refs_list[1:len(nms_and_refs_list):2]

    for nm_list, ref_list in zip(nms_list, refs_list):
        for nm, ref in zip(nm_list, ref_list):
            named_struct_list.append(NameStruct(nm, ref))

    return named_struct_list


def get_named_struct_list_from_existing_nms_only(nms_list: list, bk: xw.Book):
    """Function to convert a list of known named range names to a named struct.

    Arguments:
        nms_list {list} -- list of known names.
        bk {xw.Book} -- book names exist in.
    """
    return [NameStruct(nm, bk.names(nm).refers_to) for nm in nms_list]


def seperate_list(all_list, pred):
    """Given a predicate, seperate a list into a true and false list.

    Arguments:
        all_list {list} -- all items
        pred {function} -- predicate function
    """

    true_list = []
    false_list = []

    for item in all_list:
        if pred(item):
            true_list.append(item)
        else:
            false_list.append(item)

    return true_list, false_list


def name_not_in_all_pred_factory(all_nms):
    def inner_func(item):
        return item.nm not in all_nms
    return inner_func


def potential_ref_not_in_all_refs_factory(all_nms_structs):
    all_refs = [all_nm_struct.refers_to for all_nm_struct in all_nms_structs]

    def inner_func(item):
        return item.refers_to not in all_refs
    return inner_func


def name_needs_rename_pred_factory(tracked_nm_structs):
    """Predicate factory to determine which potential name struct
    has is not currently in the tracked name structs but its 
    potential reference is.



    Arguments:
        tracked_nm_structs {[type]} -- list of tracked name structs.

    Returns:
        tuple(list, list) -- list of names that need a rename, list of names that need a swap.
    """

    nms_list = [nm_struct.nm for nm_struct in tracked_nm_structs]
    refs_list = [nm_struct.refers_to for nm_struct in tracked_nm_structs]

    def inner_func(item):
        return item.nm not in nms_list and item.refers_to in refs_list

    return inner_func


def name_needs_no_action_pred_factory(tracked_name_structs: list):
    def inner_func(item):
        return tracked_name_structs.count(item) == 1
    return inner_func


def handle_rename(potential_struct, tracked_structs, bk):
    for tracked_struct in tracked_structs:
        if potential_struct.refers_to == tracked_struct.refers_to:
            bk.names(tracked_struct.nm).name = potential_struct.nm
            return


def handle_all_renames(rename_structs, tracked_structs, bk):
    for rename_struct in rename_structs:
        handle_rename(rename_struct, tracked_structs, bk)


def handle_all_swaps(swap_structs, tracked_structs, bk):
    swap_refs = [swap_struct.refers_to for swap_struct in swap_structs]

    for tracked_struct in tracked_structs:
        if tracked_struct.refers_to in swap_refs:
            bk.names(tracked_struct.nm).delete()

    for swap_struct in swap_structs:
        nm_helper.add_named_range_from_addr(
            bk, swap_struct.refers_to, swap_struct.nm)


def handle_name_structs_not_in_all(nm_structs, bk):
    for nm_struct in nm_structs:
        nm_helper.add_named_range_from_addr(
            bk, nm_struct.refers_to, nm_struct.nm
        )


class NameManager:
    """A class that is responsible for updating any named ranges
    on tracked sheets.

    """

    def __init__(self, bk: xw.Book):
        self.bk: xw.Book = bk

    def convert_nr_names_to_refers_to(self, nms):
        return [self.bk.names(nm).refers_to for nm in nms]

    def get_ref_err_nms_to_delete(self, nms):
        return nm_helper.get_list_of_ref_err_nms(self.bk)

    def get_list_of_potential_names(self, sht_nm, sht_type):
        return rng_getters.get_names_and_addresses(self.bk.sheets[sht_nm], sht_type)

    def delete_all_ref_error_nm_rngs(self):
        delete_nms = nm_helper.get_list_of_ref_err_nms(self.bk)
        for delete_nm in delete_nms:
            self.bk.names(delete_nm).delete()

    def handle_all_naming_cases(self, potential_nms_and_refs, all_nms, all_tracked_nms):

        all_tracked_nm_structs = get_named_struct_list_from_existing_nms_only(
            all_tracked_nms, self.bk)

        all_nm_structs = get_named_struct_list_from_existing_nms_only(
            all_nms, self.bk
        )
        named_structs = get_named_struct_list(potential_nms_and_refs)

        named_structs_not_in_all, named_structs_in_all = seperate_list(
            named_structs, potential_ref_not_in_all_refs_factory(all_nm_structs))

        named_structs_need_no_action, named_structs_need_action = seperate_list(
            named_structs_in_all, name_needs_no_action_pred_factory(
                all_tracked_nm_structs)
        )

        named_structs_need_rename, named_structs_need_swap = seperate_list(
            named_structs_need_action, name_needs_rename_pred_factory(
                all_tracked_nm_structs)
        )

        handle_name_structs_not_in_all(named_structs_not_in_all, self.bk)
        handle_all_renames(named_structs_need_rename,
                           all_tracked_nm_structs, self.bk)

        handle_all_swaps(named_structs_need_swap,
                         all_tracked_nm_structs, self.bk)
