from .base import Theme, ThemeColors

class DarkTheme(Theme):
    name = "dark"
    colors = ThemeColors(
        background="#111827",         # Tailwind gray-900
        surface="#1F2937",            # Tailwind gray-800
        surface_secondary="#374151",  # Tailwind gray-700
        
        text_primary="#F9FAFB",       # Tailwind gray-50
        text_secondary="#D1D5DB",     # Tailwind gray-300
        text_muted="#9CA3AF",         # Tailwind gray-400
        
        accent="#3B82F6",             # Tailwind blue-500
        
        positive="#22C55E",           # Tailwind green-500
        warning="#F59E0B",            # Tailwind amber-500
        negative="#EF4444",           # Tailwind red-500
        info="#0EA5E9",               # Tailwind sky-500
        
        input_background="#374151",   # Tailwind gray-700
        input_border="#4B5563",       # Tailwind gray-600
        
        border="#374151",             # Tailwind gray-700
    )
    table_style = "Table Style Medium 2"
