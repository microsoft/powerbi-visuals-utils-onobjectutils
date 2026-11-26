import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { load } from "js-yaml";

const workflow = load(readFileSync(new URL("../workflows/release.yml", import.meta.url), "utf8"));
const { draft, build, publish } = workflow.jobs;

test("release transitions and jobs require each preceding stage to succeed", () => {
    assert.deepEqual(workflow.on, { release: { types: ["created"] } });
    assert.deepEqual(Object.keys(workflow.jobs), ["draft", "build", "publish"]);
    assert.equal(draft.needs, undefined);
    assert.equal(build.needs, "draft");
    assert.equal(publish.needs, "build");

    for (const job of Object.values(workflow.jobs)) {
        assert.equal(job["runs-on"], "ubuntu-latest");
        assert.equal(job.if, undefined);
        assert.equal(job["continue-on-error"], undefined);
        for (const step of job.steps) {
            assert.equal(step.if, undefined);
            assert.equal(step["continue-on-error"], undefined);
        }
    }

    assert.equal(draft.steps.length, 1);
    assert.match(draft.steps[0].run, /^gh release edit "\$TAG_NAME" --draft=true --repo "\$GITHUB_REPOSITORY"$/);
    assert.equal(publish.steps.length, 3);
    assert.match(publish.steps[1].run, /^gh release upload "\$TAG_NAME" "release-artifact\/lib.zip" --repo "\$GITHUB_REPOSITORY"$/);
    assert.match(publish.steps[2].run, /^gh release edit "\$TAG_NAME" --draft=false --repo "\$GITHUB_REPOSITORY"$/);
});

test("package execution has only a read token and no persisted checkout credentials", () => {
    assert.deepEqual(workflow.permissions, {});
    assert.equal(workflow.env, undefined);
    assert.deepEqual(build.permissions, { contents: "read" });
    assert.equal(build.env, undefined);
    assert.equal(build.environment, undefined);
    assert.doesNotMatch(JSON.stringify(build), /secrets\.|GH_TOKEN|GITHUB_TOKEN/);

    const checkout = build.steps[0];
    assert.match(checkout.uses, /^actions\/checkout@[a-f0-9]{40}$/);
    assert.deepEqual(checkout.with, {
        ref: "${{ github.sha }}",
        "persist-credentials": false,
    });

    const commands = build.steps.filter(step => step.run).map(step => step.run.trim());
    assert.deepEqual(commands, [
        "npm ci",
        "npm run lint",
        "npm run build",
        "npx playwright install --with-deps chromium",
        "npm test",
        "zip -r lib.zip lib *.md LICENSE package.json package-lock.json",
    ]);
});

test("privileged jobs use environment gates and expose tokens only to gh commands", () => {
    for (const job of [draft, publish]) {
        assert.deepEqual(job.permissions, { contents: "write" });
        assert.equal(job.environment, "automated-release");
        assert.equal(job.env, undefined);
        for (const step of job.steps) {
            if (step.run) {
                assert.match(step.run, /^gh release (edit|upload) .* --repo "\$GITHUB_REPOSITORY"$/);
                assert.doesNotMatch(step.run, /[\r\n;&|`]/);
                assert.deepEqual(step.env, {
                    GH_TOKEN: "${{ secrets.GITHUB_TOKEN }}",
                    TAG_NAME: "${{ github.event.release.tag_name }}",
                });
            } else {
                assert.match(step.uses, /^actions\/download-artifact@[a-f0-9]{40}$/);
                assert.equal(step.env, undefined);
            }
        }
    }
});

test("the archive crosses runners as a named artifact without being executed", () => {
    const transfer = build.steps.at(-1);
    assert.match(transfer.uses, /^actions\/upload-artifact@[a-f0-9]{40}$/);
    assert.deepEqual(transfer.with, {
        name: "release-archive",
        path: "lib.zip",
        "if-no-files-found": "error",
    });
    const download = publish.steps[0];
    assert.match(download.uses, /^actions\/download-artifact@[a-f0-9]{40}$/);
    assert.deepEqual(download.with, {
        name: transfer.with.name,
        path: "release-artifact",
    });
});
