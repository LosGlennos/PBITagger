import {Command, OptionValues} from "commander";
import Tagger from "./tagger";

const program: Command = new Command();
program
    .version("1.0.0")
    .option("--accessToken <value>", "Token to use for Azure DevOps access.")
    .option("--organization <value>", "Organization that holds your project")
    .option("--projectName <value>", "Name of the team project.")
    .option("--releaseId <value>", "Release with the items to tag.")
    .parse(process.argv);

const options: OptionValues = program.opts();

if (!options.accessToken) {
    throw "No access token supplied.";
}

if (!options.releaseId) {
    throw "No organization name supplied.";
}

if (!options.projectName) {
    throw "No project name supplied.";
}

if (!options.releaseId) {
    throw "No release id supplied.";
}

Tagger.build(options.accessToken, options.organization)
    .then(tagger => {
        tagger.tagItemsInRelease(options.projectName, options.releaseId)
            .then(_ => console.log("Work items successfully tagged;"));
    });