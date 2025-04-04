if [ -z "$CAL_LINK_24_3_FDN" ]; then
    echo "Error: CAL_LINK_24_3_FDN environment variable is not set"
    exit 1
fi

mkdir -p ./downloaded

wget -qO ./downloaded/24.3-fdn.xlsx "$CAL_LINK_24_3_FDN"