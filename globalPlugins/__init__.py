# globalPlugins package



JAWS_ORDER = [
    "class","display","href","id",
    "role","xml-roles",
    "level","posinset","setsize",
    "expanded","fsFormField",
    "type","text-input-type",
    "multiline","maxlength",
    "autocomplete","haspopup",
    "name-from","explicit-name",
    "title","src","placeholder",
    "text-align","text-model"
]

def sort_params(params):
    ordered=[]
    for k in JAWS_ORDER:
        if k in params:
            ordered.append((k,params[k]))
    for k in params:
        if k not in [x[0] for x in ordered]:
            ordered.append((k,params[k]))
    return ordered
