from .base import Theme, ThemeColors

class LightTheme(Theme):
    name = "light"
    colors = ThemeColors(
        background="#F3F4F6",         # Tailwind gray-100
        surface="#FFFFFF",            # White
        surface_secondary="#F9FAFB",  # Tailwind gray-50

        text_primary="#111827",       # Tailwind gray-900
        text_secondary="#4B5563",     # Tailwind gray-600
        text_muted="#9CA3AF",         # Tailwind gray-400

        accent="#2563EB",             # Tailwind blue-600

        positive="#16A34A",           # Tailwind green-600
        warning="#D97706",            # Tailwind amber-600
        negative="#DC2626",           # Tailwind red-600
        info="#0284C7",               # Tailwind sky-600

        input_background="#FFFFFF",
        input_border="#D1D5DB",       # Tailwind gray-300

        border="#E5E7EB",             # Tailwind gray-200
    )
    table_style = "Table Style Light 1"
