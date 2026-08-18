if [ -z "$CAL_LINK" ]; then
    echo "Error: CAL_LINK environment variable is not set"
    exit 1
fi

mkdir -p ./downloaded

wget -qO ./downloaded/25.3_y2_sem1.xlsx "$CAL_LINK"