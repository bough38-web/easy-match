from diagnostics import collect_summary, format_summary
def main():
    print(format_summary(collect_summary()))
    print("OK")
if __name__ == "__main__":
    main()
