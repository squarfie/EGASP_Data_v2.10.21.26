from django import template
from operator import attrgetter

register = template.Library()

@register.filter
def dict_lookup(dictionary, key):
    """Safely retrieve a value from a dictionary using a key."""
    if isinstance(dictionary, dict):
        return dictionary.get(key, None)
    return None

@register.filter
def get_item(dictionary, key):
    """Retrieve a value from a dictionary using a key."""
    return dictionary.get(key, None)

@register.filter
def make_tuple(value1, value2):
    """Creates a tuple with two values."""
    return (value1, value2)

@register.simple_tag
def get_existing_value(existing_entries, entry_id, value_type):
    """
    Retrieves the existing value for a given breakpoint entry.
    """
    from apps.home.models import BreakpointsTable

    # Existing results can come from CSV imports or older records that have
    # an antibiotic code but no current BreakpointsTable M2M association.
    breakpoint = BreakpointsTable.objects.filter(pk=entry_id).first()
    is_retest = value_type.startswith("retest_")
    code_field = "ab_Retest_Abx_code" if is_retest else "ab_Abx_code"
    entry = None
    if breakpoint and breakpoint.Whonet_Abx:
        entry = existing_entries.filter(
            **{code_field + "__iexact": breakpoint.Whonet_Abx}
        ).first()
    if entry is None:
        entry = existing_entries.filter(ab_breakpoints_id__in=[entry_id]).first()

    if entry:
        if value_type == 'disk':
            return entry.ab_Disk_value 
        elif value_type == 'mic':
            return entry.ab_MIC_value
        elif value_type == 'retest_disk':
            return entry.ab_Retest_DiskValue
        elif value_type == 'retest_mic':
            return entry.ab_Retest_MICValue
        elif value_type == 'mic_operand':
            return entry.ab_MIC_operand or ''
        elif value_type == 'retest_mic_operand':
            return entry.ab_Retest_MIC_operand or ''
        elif value_type == 'alert_mic':
            return entry.ab_AlertMIC
        elif value_type == 'retest_alert_mic':
            return entry.ab_Retest_AlertMIC
    return ''

@register.filter
def multi_sort(queryset, fields):
    fields = fields.split(',')
    return sorted(queryset, key=attrgetter(*fields))



@register.filter
def mic_format(value):
    try:
        val = float(value)

        if val > 0.25:
            return f"{val:.1f}"   # 1 decimal place
        else:
            return f"{val:.3f}".rstrip('0').rstrip('.')  # clean trailing zeros

    except (TypeError, ValueError):
        return value