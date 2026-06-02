import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";

export function createProjectStore(projectsFile) {
  return {
    async load() {
      if (!existsSync(projectsFile)) return { projects: [] };
      const store = JSON.parse(await readFile(projectsFile, "utf8"));
      return {
        projects: Array.isArray(store.projects) ? store.projects : [],
      };
    },
    async save(store) {
      await mkdir(dirname(projectsFile), { recursive: true });
      await writeFile(
        projectsFile,
        `${JSON.stringify({ projects: store.projects ?? [] }, null, 2)}\n`,
      );
    },
  };
}
