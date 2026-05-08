#!/usr/bin/env bash
# Check laptop temperatures using lm-sensors and/or kernel thermal zones

if command -v sensors &>/dev/null; then
    sensors
elif [ -d /sys/class/thermal ]; then
    echo "lm-sensors not found — reading kernel thermal zones"
    for zone in /sys/class/thermal/thermal_zone*; do
        type=$(cat "$zone/type" 2>/dev/null)
        temp_raw=$(cat "$zone/temp" 2>/dev/null)
        if [ -n "$temp_raw" ]; then
            temp=$(echo "scale=1; $temp_raw / 1000" | bc)
            printf "%-30s %s°C\n" "$type" "$temp"
        fi
    done
else
    echo "No temperature sources found. Install lm-sensors:"
    echo "  sudo apt install lm-sensors && sudo sensors-detect"
    exit 1
fi
