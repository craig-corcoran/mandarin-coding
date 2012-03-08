import mandarin
import plac

@plac.annotations(
    spec_file_path = ("path to file containing specifications for table generation",),
    )
def main(spec_file_path):
    mandarin.generate_table(spec_file_path)

if __name__ == "__main__":
    plac.call(main)
